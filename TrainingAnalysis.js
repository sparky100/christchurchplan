
function getLaps(activityId, accessToken) {
  const url = `https://www.strava.com/api/v3/activities/${activityId}/laps`;

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: "Bearer " + accessToken
    },
    muteHttpExceptions: true
  });

  return JSON.parse(response.getContentText());
}

function writeLapsToSheet(activity, laps) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("LapData") || ss.insertSheet("LapData");

  // Header row if empty
  if (sheet.getLastRow() === 0) {
    const header = [
      "ActivityID",
      "ActivityName",
      "ActivityDate",
      "LapIndex",
      "LapStartTime",
      "ElapsedTime",
      "MovingTime",
      "Distance",
      "AvgSpeed",
      "MaxSpeed",
      "AvgHR",
      "MaxHR",
      "ElevationGain"
    ];
    sheet.appendRow(header);
  }

  // Build a set of existing unique keys: ActivityID|LapIndex
  const lastRow = sheet.getLastRow();
  const existingKeys = new Set();

  if (lastRow > 1) {
    const keyRange = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    keyRange.forEach(row => {
      const key = row[0] + "|" + row[3]; // ActivityID|LapIndex
      existingKeys.add(key);
    });
  }

  // Build new rows only
  const newRows = [];

  laps.forEach(lap => {
    const key = activity.id + "|" + lap.lap_index;

    if (!existingKeys.has(key)) {
      newRows.push([
        activity.id,
        activity.name,
        activity.start_date_local,
        lap.lap_index,
        lap.start_date,
        lap.elapsed_time,
        lap.moving_time,
        lap.distance,
        lap.average_speed,
        lap.max_speed,
        lap.average_heartrate,
        lap.max_heartrate,
        lap.total_elevation_gain
      ]);
    }
  });

  // Batch write new rows
  if (newRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }
}


function getISOWeek(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Refresh Strava access token using refresh token.
 */


function refreshStravaToken() {
  const clientId = '203400';
  const clientSecret = '5a36e68f3c42c63ddb3f5e6194c55e9d176fa3c7';
  const refreshToken = '37485e330e260f4105f5d5cf8d1955ef05d8bad5';

  const url = 'https://www.strava.com/oauth/token';

  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: 'refresh_token'
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (!data.access_token) {
    throw new Error("Failed to refresh Strava token: " + response.getContentText());
  }

  PropertiesService.getScriptProperties().setProperty('STRAVA_ACCESS_TOKEN', data.access_token);
  return data.access_token;
}

function fetchActivityDetail(id) {
  let token = PropertiesService.getScriptProperties().getProperty('STRAVA_ACCESS_TOKEN');

  const url = `https://www.strava.com/api/v3/activities/${id}?include_all_efforts=true`;

  let response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  // If token expired, refresh and retry once
  if (response.getResponseCode() === 401) {
    token = refreshAccessToken();

    response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
  }

  return JSON.parse(response.getContentText());
}



/**
 * Import Strava activities, dedupe, append, notify.
 */
function importStravaActivities() {
  const spreadsheetId = '1NL4pgFfkdqEvbm45EDw-nR3O9eQbFgdwvrGhgZ6sqKU';
  const sheetName = 'StravaActivities';
  const email = 'parkinsonsj@gmail.com';

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

  // Get access token
  let accessToken = PropertiesService.getScriptProperties().getProperty('STRAVA_ACCESS_TOKEN');
  if (!accessToken) accessToken = refreshStravaToken();

  const url = 'https://www.strava.com/api/v3/athlete/activities?per_page=50';
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch(url, options);

  // Retry if token expired
  if (response.getResponseCode() === 401) {
    accessToken = refreshStravaToken();
    options.headers.Authorization = 'Bearer ' + accessToken;
    response = UrlFetchApp.fetch(url, options);
  }

  // Strava should return an array here 
  const body = response.getContentText(); 
  Logger.log(body); 
  const activities = JSON.parse(body);
  // Strava should return an array here 
  if (!Array.isArray(activities)) 
  { 
    throw new Error('Expected an array of activities but got: ' + JSON.stringify(activities)); 
  }

  // Existing dup
  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String)
    : [];

  const existingSet = new Set(existingIds);
  const newRows = [];
  const planStart = new Date('2026-02-16'); // the start of the training plan
  const zoneRows = [];



activities
  .filter(a =>
    (a.type === "Run" || a.type === "TrailRun" || a.type === "VirtualRun") &&
    new Date(a.start_date_local) >= planStart
  )
  .forEach(a => {
    if (!existingSet.has(String(a.id))) {

      // ---- Your existing calculations ----
      const distanceKm = +(a.distance / 1000).toFixed(2); 
      const movingMin = +(a.moving_time / 60).toFixed(1);
      const elaspsedMin = +(a.elapsed_time / 60).toFixed(1);

      const Rest_HR = 54
      const Max_HR = 176

      const hrReserve = (a.average_heartrate - Rest_HR) / (Max_HR - Rest_HR);
      const pace = movingMin / distanceKm;
      let efficiency =pace / hrReserve;
      let load = movingMin * (1 / efficiency);

      const activityDate = new Date(a.start_date_local); 
      const dateOnly = new Date(activityDate.getFullYear(), activityDate.getMonth(), activityDate.getDate());

      const isoWeek = getISOWeek(dateOnly);
            // ---- NEW: Fetch HR + pace zone data ----
      // Inside your per-activity loop
      const { hrTimes, paceTimes } = getActivityZoneData(a.id);

      let laps = getLaps(a.id, accessToken);
      writeLapsToSheet(a, laps);

newRows.push([
  a.id,
  dateOnly,
  a.name,
  distanceKm,
  movingMin,
  elaspsedMin,
  pace,
  a.type,
  a.average_heartrate,
  a.max_heartrate,
  ...hrTimes,     // 5 HR zone columns
  ...paceTimes,    // 6 pace zone columns
  isoWeek,
  efficiency,
  load
]);

    }
  });

Logger.log(newRows)
  // Append new rows
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
         .setValues(newRows);


    // Send notification
    MailApp.sendEmail({
      to: email,
      subject: `Strava Import: ${newRows.length} new activities added`,
      htmlBody: `
        <p>Hi Simon,</p>
        <p>${newRows.length} new Strava activities were imported into your sheet:</p>
        <ul>
          ${newRows.map(r => `<li>${r[2]} — ${r[3].toFixed(2)} km</li>`).join('')}
        </ul>
        <p>Everything is deduplicated and up to date.</p>
      `
    });
  }
}

function fetchActivityZones(activityId) {
  const token = PropertiesService.getScriptProperties().getProperty('STRAVA_ACCESS_TOKEN');
  const url = `https://www.strava.com/api/v3/activities/${activityId}/zones`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 401) {
    // Token expired → refresh once
    const newToken = refreshAccessToken();
    const retry = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${newToken}` },
      muteHttpExceptions: true
    });
    return JSON.parse(retry.getContentText());
  }

  return JSON.parse(response.getContentText());
}

function extractZoneTimes(zonesJson) {
  let hrTimes = null;
  let paceTimes = null;

  zonesJson.forEach(z => {
    if (z.type === "heartrate" && Array.isArray(z.distribution_buckets)) {
      hrTimes = z.distribution_buckets.map(b => +(b.time / 60).toFixed(1));
    }
    if (z.type === "pace" && Array.isArray(z.distribution_buckets)) {
      paceTimes = z.distribution_buckets.map(b => +(b.time / 60).toFixed(1));
    }
  });

  if (!hrTimes) hrTimes = [0, 0, 0, 0, 0];
  if (!paceTimes) paceTimes = [0, 0, 0, 0, 0, 0];

  return { hrTimes, paceTimes };
}


function getActivityZoneData(activityId) {
  const zonesJson = fetchActivityZones(activityId);
  return extractZoneTimes(zonesJson);
}


function testActivityZoneData() {
  const activityId = 17458077685; // change to any activity ID you like

  const { hrTimes, paceTimes } = getActivityZoneData(activityId);

  console.log("HR zone times (sec):", hrTimes);
  console.log("Pace zone times (sec):", paceTimes);
}


/**
 * Run this once per day (you will add the trigger manually in the UI).
 * Builds the analysis object, calls AI, sends plain-text email.
 */
function runDailyTrainingAnalysis() {
  const todaysActivities = getTodaysActivities_();
  if (todaysActivities.length === 0) {
    // No activities today – nothing to analyse.
    return;
  }
  Logger.log(todaysActivities)
  const todaysLaps = getLapsForActivities_(todaysActivities);
  const last14DaysActivities = getLast14DaysActivities_();
  const keyStats = getKeyStats_(); // { ctl, atl, tsb }

  Logger.log(todaysLaps)
  Logger.log("Last 14 days")
  Logger.log(last14DaysActivities)

  Logger.log(keyStats)

  const analysisData = {
    target: {
      race: "Half Marathon",
      date: "2026-04-12",
      targetTime: "1:34:00",
      targetPace: "4:28/km"
    },
    todaysActivities: todaysActivities,
    todaysLaps: todaysLaps,
    last14Days: {
      activities: last14DaysActivities,
      ctl: keyStats.ctl,
      atl: keyStats.atl,
      tsb: keyStats.tsb
    }
  };

  const payload = { todaysActivities,todaysLaps, last14DaysActivities, keyStats, analysisData };

  const analysis = sendDataToClaudeHaiku(payload);

  saveMarkdownToDrive(analysis);

  const mdBlob = Utilities.newBlob(
  analysis,
  'text/markdown',
  'DailyTrainingAnalysis_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") + ".md"
);

GmailApp.sendEmail(
  Session.getActiveUser().getEmail(),
   "Daily Training Analysis",
  "Please find your daily training analysis attached.",
  { attachments: [mdBlob] }
);

}


/**
 * Reads today's activities from StravaActivities using header-based mapping.
 */
function getTodaysActivities_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("StravaActivities");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const rows = values.slice(1);

  const todayStr = new Date().toDateString();

  return rows
    .map(row => Object.fromEntries(header.map((h, i) => [h, row[i]])))
    .filter(a => {
      const d = new Date(a.Date);
      return !isNaN(d) && d.toDateString() === todayStr;
    });
}

/**
 * Reads all laps for the given activities from LapData, keyed by ActivityID.
 */
function getLapsForActivities_(activityList) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("LapData");
  const lapMap = {};

  Logger.log(activityList)
  if (!sheet || activityList.length === 0) return lapMap;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return lapMap;

  const header = values[0];
  const rows = values.slice(1);

  activityList.forEach(a => {
    lapMap[a.ActivityID] = [];
  });

  rows.forEach(row => {
    const lap = Object.fromEntries(header.map((h, i) => [h, row[i]]));
    const id = lap.ActivityID;
    if (id && lapMap.hasOwnProperty(id)) {
      lapMap[id].push(lap);
    }
  });

  return lapMap;
}

/**
 * Reads all activities from the last 14 days from StravaActivities.
 */
function getLast14DaysActivities_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("StravaActivities");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const rows = values.slice(1);

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 14);
  Logger.log("Pre-filter")
  Logger.log(rows)
  Logger.log(cutoff)


  const filtered =rows
    .map(row => Object.fromEntries(header.map((h, i) => [h, row[i]])))
    .filter(a => {
     const d = new Date(a.Date);
     return d >= cutoff;
    });
Logger.log("filtered")
 Logger.log(filtered)

return filtered   
 
  
}

/**
 * Reads current CTL, ATL, TSB from KeyStats!B11, C11, D11.
 */
function getKeyStats_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Key Stats");
  if (!sheet) {
    return { ctl: null, atl: null, tsb: null };
  }

  const ctl = sheet.getRange("B11").getValue(); // Fitness (CTL)
  const atl = sheet.getRange("C11").getValue(); // Fatigue (ATL)
  const tsb = sheet.getRange("D11").getValue(); // Recovery (TSB)

  Logger.log(ctl,atl,tsb)

  return { ctl: ctl, atl: atl, tsb: tsb };
}

function callClaude(promptText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");

  const url = "https://api.anthropic.com/v1/messages";

  const payload = {
    model: "claude-3-haiku-20240307",
    max_tokens: 2000,
    messages: [
      { role: "user", content: promptText }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  return json.content?.[0]?.text || "No response";
}

function sendDataToClaudeHaiku(payload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");

  const url = "https://api.anthropic.com/v1/messages";

  const prompt = `
You are an expert running coach. Analyse the following dataset:

${JSON.stringify(payload, null, 2)}

I am calculating CTL (Chronic Training Load) for running using the following formula:
CTL = sum of Load scores over the last 42 days / MIN(42, days since first activity)
Where Load for each activity is calculated as:
load = movingMin * (1 / efficiency)
efficiency = pace / hrReserve
hrReserve = (avgHR - restHR) / (maxHR - restHR)
pace = movingMin / distanceKm
Personal constants: restHR = [x], maxHR = [x]
ATL uses the same load metric over 7 days. TSB = CTL - ATL.
The RAG thresholds for this metric are:

CTL: Red <3, Amber 3–5, Green 5–8, Blue 8+
ATL: Green <3, Amber 3–8, Red 8–12
TSB: Red <−6, Amber −6 to −2, Green −2 to +3, Amber >+3

When analysing my training data, always interpret CTL, ATL and TSB values in the context of these specific formulas and thresholds, not generic TrainingPeaks-style TSS values which operate on a different scale'

Multiple activities on the same day may form part of the same session. 

I am based in the UK. The running week is Monday to Sunday. Make sure that you correctlty identify the current UK dates when suggesting next actions.

Provide:
- a summary of the last 14 days of training
- load, fatigue, and freshness interpretation
- pacing and intensity insights
- whether the training aligns with a target half marathon time of 1:34
- provide a probality assessment of predicted times
- actionable recommendations for the next 48 hours - This is important

`;

  const requestBody = {
    model: "claude-opus-4-5",
    max_tokens: 4096,
    messages: [
      { role: "user", content: prompt }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  return json.content?.[0]?.text || "No response";
}

function testClaudeConnection() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");

  const url = "https://api.anthropic.com/v1/messages";

  const payload = {
    model: "claude-opus-4-5",
    max_tokens: 50,
    messages: [
      { role: "user", content: "Reply with the word: SUCCESS" }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  Logger.log(json);

  return json.content?.[0]?.text || "No response";
}

function saveMarkdownToDrive(markdownText) {
  const folder = DriveApp.getFolderById("1zjch0T0aYihHxG_WaVhP1p6N7Ceg2JdC"); // or DriveApp.getRootFolder()
  const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  const filename = `training-analysis-${timestamp}.md`;

  folder.createFile(filename, markdownText, MimeType.PLAIN_TEXT);
}

