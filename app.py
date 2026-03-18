import os
import argparse
import logging
from datetime import datetime, date as date_type
from dotenv import load_dotenv
from garminconnect import Garmin
from garminconnect.workout import (
    RunningWorkout,
    WorkoutSegment,
    ExecutableStep,
    StepType,
    ConditionType,
    TargetType,
)

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)


# ─── GARMIN AUTH ──────────────────────────────────────────────

def get_garmin_client():
    email  = os.environ['GARMIN_EMAIL']
    passwd = os.environ['GARMIN_PASSWORD']
    log.info(f"Logging in as {email}...")
    client = Garmin(email, passwd)
    client.login()
    log.info("Logged in OK")
    return client


# ─── GARMIN SCHEDULE / DELETE ─────────────────────────────────

def schedule_workout(client, workout_id, date):
    return client.connectapi(
        f'/workout-service/schedule/{workout_id}',
        method='POST',
        json={'date': date}
    )

def delete_workout(client, workout_id):
    return client.connectapi(
        f'/workout-service/workout/{workout_id}',
        method='DELETE'
    )


# ─── PACE PARSING ─────────────────────────────────────────────

def parse_pace_value(val):
    """Convert 'MM:SS' pace string to decimal min/km. e.g. '4:30' -> 4.5"""
    val = str(val).strip()
    if ':' in val:
        parts = val.split(':')
        return int(parts[0]) + int(parts[1]) / 60
    return float(val)


def parse_pace(val):
    """
    Parse pace column into (low, high) tuple in decimal min/km.

    Handles:
        '4:00-4:30'  -> (4.0, 4.5)      range
        '4:30'       -> (4.5, 4.5)      single — +/-10s tolerance added in pace_target()
        '6:00+'      -> (6.0, 99.0)     slower than 6:00, no upper limit
        'N/A'        -> None
        ''           -> None
    """
    if not val or str(val).strip().upper() == 'N/A':
        return None

    val = str(val).strip()
    val = val.replace('\u2013', '-').replace('\u2014', '-')  # normalise en/em dashes

    if val.endswith('+'):
        low = parse_pace_value(val[:-1].strip())
        return (low, 99.0)

    if '-' in val:
        parts = val.split('-')
        low   = parse_pace_value(parts[0].strip())
        high  = parse_pace_value(parts[1].strip())
        return (low, high)

    p = parse_pace_value(val)
    return (p, p)


# ─── TARGETS ──────────────────────────────────────────────────

def pace_target(pace_range):
    """
    Build a Garmin pace zone target from a (low, high) min/km tuple.
    Garmin stores pace as m/s:
      targetValueOne = faster bound (higher m/s)
      targetValueTwo = slower bound (lower m/s)
    """
    low, high = pace_range

    if low == high:
        tolerance = 10 / 60  # +/-10 secs/km
        low  = low  - tolerance
        high = high + tolerance

    fast_ms = 1000 / (low  * 60)  # faster bound in m/s
    slow_ms = 1000 / (high * 60)  # slower bound in m/s

    return {
        'workoutTargetTypeId':  6,
        'workoutTargetTypeKey': 'pace.zone',
        'displayOrder':         6,
        'targetValueOne':       fast_ms,
        'targetValueTwo':       slow_ms,
    }

def no_target():
    return {
        'workoutTargetTypeId':  1,
        'workoutTargetTypeKey': 'no.target',
        'displayOrder':         1,
    }


# ─── DISTANCE STEP ────────────────────────────────────────────

def create_distance_step(distance_meters, step_order, target=None):
    """Distance-based step with pace target at step level."""
    t = dict(target or no_target())

    # Extract target values — must be at step level not inside targetType
    target_value_one = t.pop('targetValueOne', None)
    target_value_two = t.pop('targetValueTwo', None)

    step = ExecutableStep(
        stepOrder=step_order,
        stepType={
            'stepTypeId':   StepType.INTERVAL,
            'stepTypeKey':  'interval',
            'displayOrder': 3,
        },
        endCondition={
            'conditionTypeId':  3,
            'conditionTypeKey': 'distance',
            'displayOrder':     3,
            'displayable':      True,
        },
        endConditionValue=float(distance_meters),
        targetType=t
    )

    if target_value_one is not None:
        step.targetValueOne = target_value_one
    if target_value_two is not None:
        step.targetValueTwo = target_value_two

    return step


# ─── ROW TO WORKOUT ───────────────────────────────────────────

def row_to_workout(row):
    """
    Convert a sheet row to a Garmin RunningWorkout.

    Sheet columns used:
        Description      — workout name
        Plan_Dist        — distance in km
        Pace (min/km)    — pace target: '4:00-4:30', '4:30', '6:00+', or 'N/A'
    """
    name = str(row.get('Full_description', '')).strip()
    distance = float(row.get('Plan_Dist') or 0) * 1000  # km -> meters
    pace_raw = row.get('Pace (min/km)', '')
    pace     = parse_pace(pace_raw)
    target   = pace_target(pace) if pace else no_target()

    steps = [create_distance_step(distance, 1, target)]

    segment = WorkoutSegment(
        segmentOrder=1,
        sportType={'sportTypeId': 1, 'sportTypeKey': 'running'},
        workoutSteps=steps
    )

    # Estimate duration from pace midpoint x distance
    if pace and distance:
        low, high      = pace
        mid_pace       = (low + high) / 2 if high < 90 else low
        estimated_secs = int((distance / 1000) * mid_pace * 60)
    else:
        estimated_secs = 3600

    return RunningWorkout(
        workoutName=name,
        estimatedDurationInSecs=estimated_secs,
        workoutSegments=[segment]
    )


# ─── ACTIONS ──────────────────────────────────────────────────

def action_test(client):
    print("Fetching workouts...", flush=True)
    workouts = client.get_workouts(0, 10)
    print(f"Found {len(workouts)} workouts\n", flush=True)
    print(f"{'ID':<15} Name", flush=True)
    print('-' * 50, flush=True)
    for w in workouts:
        print(f"{w['workoutId']:<15} {w['workoutName']}", flush=True)
    print(flush=True)


def action_push(client, worksheet):
    today   = date_type.today()
    rows    = worksheet.get_all_records()
    pushed  = 0
    skipped = 0
    errors  = 0

    for i, row in enumerate(rows, start=2):
        date      = str(row.get('Date', '')).strip()
        name = str(row.get('Full_description', '')).strip()
        distance  = float(row.get('Plan_Dist') or 0)
        garmin_id = str(row.get('GARMIN_ID', '')).strip()

        # Skip rows with no date, name or distance
        if not date or not name or not distance:
            skipped += 1
            continue

        # Skip already pushed rows
        if garmin_id and not garmin_id.startswith('ERROR') and garmin_id != 'DUPLICATE':
            log.info(f"Row {i} — skipping '{name}' (already pushed: {garmin_id})")
            skipped += 1
            continue

        # Parse and normalise date
        parsed_date = None
        for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'):
            try:
                parsed_date = datetime.strptime(date, fmt).date()
                date        = parsed_date.strftime('%Y-%m-%d')
                break
            except ValueError:
                continue

        # Skip past dates
        if parsed_date and parsed_date < today:
            log.info(f"Row {i} — skipping '{name}' (past date {date})")
            skipped += 1
            continue

        try:
            workout    = row_to_workout(row)
            log.info(f"Row {i} — uploading '{name}' {distance}km on {date}...")
            result     = client.upload_running_workout(workout)
            workout_id = str(result['workoutId'])
            log.info(f"Row {i} — created {workout_id}")

            schedule_workout(client, workout_id, date)
            log.info(f"Row {i} — scheduled on {date} ✅")

            worksheet.update_cell(i, 15, workout_id)  # col O
            pushed += 1

        except Exception as e:
            import traceback
            log.error(f"Row {i} — failed: {e}")
            log.error(traceback.format_exc())
            worksheet.update_cell(i, 15, f'ERROR: {e}')
            errors += 1

    print(f"\nDone — pushed: {pushed}, skipped: {skipped}, errors: {errors}\n", flush=True)


def action_delete(client, worksheet):
    rows    = worksheet.get_all_records()
    confirm = input("Delete all pushed workouts from Garmin Connect? (yes/no): ")
    if confirm.lower() != 'yes':
        print("Aborted.")
        return

    deleted = 0
    for i, row in enumerate(rows, start=2):
        garmin_id = str(row.get('GARMIN_ID', '')).strip()
        if not garmin_id or garmin_id.startswith('ERROR') or garmin_id == 'DUPLICATE':
            continue
        try:
            delete_workout(client, garmin_id)
            worksheet.update_cell(i, 15, '')  # col O
            log.info(f"Deleted {garmin_id} — {row.get('Description')}")
            deleted += 1
        except Exception as e:
            log.error(f"Could not delete {garmin_id}: {e}")

    print(f"\nDeleted {deleted} workouts\n", flush=True)


def get_sheet():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_file(
        os.environ.get('GOOGLE_CREDENTIALS', 'credentials.json'),
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
    )
    gc    = gspread.authorize(creds)
    sheet = gc.open_by_key(os.environ['GOOGLE_SHEET_ID'])
    return sheet.worksheet(os.environ.get('GOOGLE_SHEET_NAME', 'TrainingPlan'))


# ─── ENTRY POINT ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Garmin training plan pusher')
    group  = parser.add_mutually_exclusive_group()
    group.add_argument('--test',   action='store_true', help='Test Garmin connection only')
    group.add_argument('--push',   action='store_true', help='Push plan to Garmin calendar')
    group.add_argument('--delete', action='store_true', help='Delete all pushed workouts')
    args = parser.parse_args()

    client = get_garmin_client()

    if args.test:
        action_test(client)
    elif args.push:
        action_push(client, get_sheet())
    elif args.delete:
        action_delete(client, get_sheet())
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
