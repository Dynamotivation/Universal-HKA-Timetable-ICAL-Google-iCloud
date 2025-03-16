/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Make a copy:
*      New Interface: Go to the project overview icon on the left (looks like this: ⓘ), then click the "copy" icon on the top right (looks like two files on top of each other)
*      Old Interface: Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Settings: Change lines 24-39 to be the settings that you want to use
* 3) Install:
*      New Interface: Make sure your toolbar says "install" to the right of "Debug", then click "Run"
*      Old Interface: Click "Run" > "Run function" > "install"
* 4) Authorize: You will be prompted to authorize the program and will need to click "Advanced" > "Go to Universal-HKA-Timetable-ICAL-Google-iCloud (unsafe)"
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
* 5) You can also run "startSync" if you want to sync only once (New Interface: change the dropdown to the right of "Debug" from "install" to "startSync")
*
* **To stop the Script from running click in the menu "Run" > "Run function" > "uninstall" (New Interface: change the dropdown to the right of "Debug" from "install" to "uninstall")
*
*=========================================
*               SETTINGS
*=========================================
*/

const runIntervalMinutes = 15;        // How often to run scrape the time table in minutes
const calendarName = "HKA Timetable"; // Name of the Google Calendar to add events to
const weeksToScan = 4;                // How many weeks to scan into the future not including current week
const blacklist = [                   // Blacklist entire event series by blacklisting the name or unique events by blacklisting the exclusion hash from the description

];
const coursesToCheck = [              // Identifiers of which data to fetch. Use developer mode to find this out on https://raumzeit.hka-iwi.de/timetables
  {
    "courseOfStudy": "MINB",
    "faculty": "IWI",
    "semester": "MINB.2"
  },
  {
    "courseOfStudy": "MINB",
    "faculty": "IWI",
    "semester": "MINB.1"
  }
];

/*
*=========================================
*           ABOUT THE AUTHOR
*=========================================
*
* This program was created by Dynamotivation
*
* https://github.com/dynamotivation
* https://dynamotivation.com
*
*
* Inspired by GAS-ICS-Sync by Derek Antrican
*
* https://github.com/derekantrican
*
*=========================================
*            BUGS/FEATURES
*=========================================
*
* Please report any issues at https://github.com/Dynamotivation/Universal-HKA-Timetable-ICAL-Google-iCloud/issues
*
*=========================================
*            DO NOT TOUCH BELOW
*=========================================
*/

const nonceRegex = /nonce="([A-Za-z0-9]+)"/;
const viewStateRegex = /name="jakarta\.faces\.ViewState".*?value="([-\d]+:[-\d]+)"/;
const jSessionIdRegex = /JSESSIONID=([^;]+)/;
const eventsRegex = /<!\[CDATA\[(.*?)\]\]>/;

function install(){
  deleteAllTriggers();

  //Schedule sync routine to explicitly repeat and schedule the initial sync
  ScriptApp.newTrigger("startSync").timeBased().everyMinutes(getValidTriggerFrequency(runIntervalMinutes)).create();
  ScriptApp.newTrigger("startSync").timeBased().after(1000).create();
}

function startSync() {
  Logger.log("Executing Universal-HKA-Timetable-ICAL-Google-iCloud Version 1.0");

  var events = [];

  for (var i = 0; i < coursesToCheck.length; i++) {
    var currentEvents = scanTimetableForCourse(coursesToCheck[i]);
    events = events.concat(currentEvents);
  }

  eventsToGoogleCalendar(events);
}

function scanTimetableForCourse(courseToCheck) {
  var events = [];

  // Build request
  var url = "https://raumzeit.hka-iwi.de/timetables";

  var headers = {
    "Accept": 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    "Accept-Encoding": 'gzip, deflate, br, zstd',
    "Accept-Language": 'en-GB,en;q=0.9',
    "Cache-Control": 'no-cache',
    "Connection": 'keep-alive',
    // Just input any old JSESSIONID, it should signal we are returning visitor and allow us to prefill faculty, course and semester
    "Cookie": 'JSESSIONID=BA9725A32FC9076DA669E5C0F379BC13; raumzeit-courseofstudy=' + courseToCheck.courseOfStudy + '; raumzeit-faculty=' + courseToCheck.faculty + '; raumzeit-semester=' + courseToCheck.semester,
    "DNT": '1',
    "Pragma": 'no-cache',
    "Sec-Fetch-Dest": 'document',
    "Sec-Fetch-Mode": 'navigate',
    "Sec-Fetch-Site": 'none',
    "Sec-Fetch-User": '?1',
    "Upgrade-Insecure-Requests": '1',
    "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
    "sec-ch-ua": '"Not A(Brand";v="8", "Chromium";v="132", "Microsoft Edge";v="132"',
    "sec-ch-ua-mobile": '?0',
    "sec-ch-ua-platform": '"Windows"'
  };

  var options = {
    method: "get",  // Use GET request
    headers: headers,
    muteHttpExceptions: true // Optional: to prevent exceptions on error
  };


  try {
    // Send request
    var response = UrlFetchApp.fetch(url, options);  // Make the HTTP request
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();  // Get the response body as text

    Logger.log("Initial Response Code: " + responseCode);

    if (responseCode != 200) {
      Logger.log("Request unsuccessful.");
      return events;
    }

    var headers = response.getAllHeaders();

    // Extract JSessionId
    if (headers['Set-Cookie']) {
      var jSessionIdMatch;

      for (var i = 0; i < headers['Set-Cookie'].length; i++) {
        var jSessionIdMatch = headers['Set-Cookie'][i].match(jSessionIdRegex);
        if (jSessionIdMatch) break;
      }

      if (jSessionIdMatch && jSessionIdMatch[1]) {
        Logger.log("Found JSessionId: " + jSessionIdMatch[1])
      }
      else {
        Logger.log("No JSessionId found.");
        return events;
      }
    } else {
      Logger.log("No JSessionId found.");
      return events;
    }

    // Extract Nonce
    var nonceMatch = responseBody.match(nonceRegex);

    if (nonceMatch && nonceMatch[1]) {
      Logger.log("Found nonce: " + nonceMatch[1]);
    } else {
      Logger.log("No nonce found.");
      return events;
    }

    // Extract ViewState
    var viewStateMatch = responseBody.match(viewStateRegex);

    if (viewStateMatch && viewStateMatch[1]) {
      Logger.log("Found ViewState: " + viewStateMatch[1]);
    } else {
      Logger.log("No ViewState found.");
      return events;
    }

    var i = 0;
    do {
      // Fetch timetable using JSessionId, Nonce and ViewState
      var timeTable = fetchTimetable(jSessionIdMatch[1], nonceMatch[1], viewStateMatch[1]);

      if (timeTable) {
        Logger.log("Successfully fetched timetable");
      }
      else {
        Logger.log("No timetable received.");
        continue;
      }

      // Extract events JSON
      cdataMatch = timeTable.match(eventsRegex);

      if (cdataMatch && cdataMatch[1]) {
        var jsonString = cdataMatch[1];

        try {
          // Parse the new events data
          var newJsonData = JSON.parse(jsonString);

          // Append the new events to the existing jsonData array
          events = events.concat(newJsonData.events);
        } catch (error) {
          Logger.log("Error parsing or appending JSON: " + error);
        }
      } else {
        Logger.log("No CDATA block found.");
      }

      gotoNextWeek(jSessionIdMatch[1], nonceMatch[1], viewStateMatch[1]);
      i++;
    } while (i < weeksToScan);
  } catch (error) {
    Logger.log("Error: " + error.message);
  }

  return events;
}

function gotoNextWeek(jSessionId, nonce, viewState) {
  viewState = viewState.replace(/:/g, '%3A');

  var url = "https://raumzeit.hka-iwi.de/timetables.xhtml";

  // Simulate next week button press request
  var options = {
    "method": 'post',
    "headers": {
      'Accept': 'application/xml, text/xml, */*; q=0.01',
      'Accept-Encoding': 'gzip, deflate, br, zstd',
      'Accept-Language': 'en-GB,en;q=0.9',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'Cookie': 'JSESSIONID=' + jSessionId + '; raumzeit-faculty=IWI; raumzeit-courseofstudy=MINB; raumzeit-semester=MINB.2',
      'DNT': '1',
      'Faces-Request': 'partial/ajax',
      'Origin': 'https://raumzeit.hka-iwi.de',
      'Pragma': 'no-cache',
      'Referer': 'https://raumzeit.hka-iwi.de/timetables',
      'Sec-Fetch-Dest': 'empty',
      'Sec-Fetch-Mode': 'cors',
      'Sec-Fetch-Site': 'same-origin',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
      'X-Requested-With': 'XMLHttpRequest',
      'sec-ch-ua': '"Not A(Brand";v="8", "Chromium";v="132", "Microsoft Edge";v="132"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"Windows"'
    },
    "payload": "jakarta.faces.partial.ajax=true&jakarta.faces.source=form1%3Anext_week&jakarta.faces.partial.execute=form1%3Anext_week&jakarta.faces.partial.render=form1&jakarta.faces.behavior.event=click&jakarta.faces.partial.event=click&jakarta.faces.ViewState=" + viewState + "&primefaces.nonce=" + nonce,
    "muteHttpExceptions": true,
    "followRedirects": true
  };

  UrlFetchApp.fetch(url, options);
}

function fetchTimetable(jSessionId, nonce, viewState) {
  // Simple sanitization
  viewState = viewState.replace(/:/g, '%3A');

  var url = "https://raumzeit.hka-iwi.de/timetables.xhtml";

  // Get actual event data
  var options = {
    "method": "post",
    "headers": {
      "Accept": 'application/xml, text/xml, */*; q=0.01',
      "Accept-Encoding": 'gzip, deflate, br, zstd',
      "Accept-Language": 'en-GB,en;q=0.9',
      "Cache-Control": 'no-cache',
      "Connection": 'keep-alive',
      "Content-Type": 'application/x-www-form-urlencoded; charset=UTF-8',
      "Cookie": 'JSESSIONID=' + jSessionId + '; raumzeit-courseofstudy=MINB; raumzeit-faculty=IWI; raumzeit-semester=MINB.2',
      "DNT": '1',
      "Faces-Request": 'partial/ajax',
      "Origin": 'https://raumzeit.hka-iwi.de',
      "Pragma": 'no-cache',
      "Referer": 'https://raumzeit.hka-iwi.de/timetables',
      "Sec-Fetch-Dest": 'empty',
      "Sec-Fetch-Mode": 'cors',
      "Sec-Fetch-Site": 'same-origin',
      "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
      "X-Requested-With": 'XMLHttpRequest',
      "sec-ch-ua": '"Not A(Brand";v="8", "Chromium";v="132", "Microsoft Edge";v="132"',
      "sec-ch-ua-mobile": "?0",
      "sec-ch-ua-platform": '"Windows"',
      "accept": "application/xml, text/xml, */*; q=0.01",
    },
    // Basically everything in this payload does not matter beyond simply existing
    // The start and end times do absolutely nothing, hence we have the simulated gotoNextWeek method.
    "payload": "jakarta.faces.partial.ajax=true&jakarta.faces.source=form1%3Atimetable&jakarta.faces.partial.execute=form1%3Atimetable&jakarta.faces.partial.render=form1%3Atimetable&form1%3Atimetable=form1%3Atimetable&form1%3Atimetable_event=true&form1%3Atimetable_start=2025-03-16T23%3A00%3A00.000Z&form1%3Atimetable_end=2025-03-21T23%3A00%3A00.000Z&form1%3Atimetable_view=timeGridWeek&jakarta.faces.ViewState=" + viewState.replace(/:/g, '%3A') + "&primefaces.nonce=" + nonce,
    "followRedirects": true,
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    Logger.log("Timetable Response Code: " + responseCode);
    return responseBody;
  } catch (error) {
    Logger.log("Error: " + error.message);
  }

  return null;
}

function eventsToGoogleCalendar(events) {
  // Generate a unique event for each calendar entry so we can do blacklisting
  events.forEach(event => {
    event.hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, event.title + event.start + event.end)
      .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2))
      .join('');
  });

  // Mangle the data into a more usable format
  // Grouped by dates as object keys, sorted from recent to latest per day
  var groupedEvents = events
    .sort(function (a, b) {
      var dateA = new Date(a.start);
      var dateB = new Date(b.start);
      return dateA - dateB;
    })
    .reduce(function (groups, event) {
      // Extract date as "YYYY-MM-DD"
      var eventDate = new Date(event.start).toISOString().split('T')[0];

      // Add the event to the corresponding group (date)
      if (!groups[eventDate]) {
        groups[eventDate] = [];
      }
      groups[eventDate].push(event);

      return groups;
    }, {});

  // Calculate the scan begin and end dates
  var startDate = new Date(); // Today
  var daysUntilSaturday = 6 - startDate.getDay();

  var endDate = new Date();
  endDate.setDate(startDate.getDate() + daysUntilSaturday + 7 * weeksToScan);

  // Get the users HKA timetable calendar
  var calendars = CalendarApp.getAllCalendars();
  var targetCalendar = null;

  // Get or create calendar
  for (var i = 0; i < calendars.length; i++) {
    if (calendars[i].getName() === calendarName) {
      targetCalendar = calendars[i];
      break;
    }
  }

  if (!targetCalendar) {
    targetCalendar = CalendarApp.createCalendar(calendarName);
  }

  Logger.log("Got calendar");

  // Go day by day purging all events before adding newly scanned events.
  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    var targetDate = new Date(d);
    targetDate.setHours(0, 0, 0, 0);

    Logger.log("Deleting old events on: " + targetDate)

    var endOfDay = new Date(targetDate);
    endOfDay.setHours(23, 59, 59, 999);

    // Fetch all events for the given day
    var events = targetCalendar.getEvents(targetDate, endOfDay);


    // Loop through events and delete those that don't have the test tag
    events.forEach(function (event) {
      retryWithBackoff(function () {
        var tags = event.getAllTagKeys();
        if (tags.includes("hka_scraper")) {
          event.deleteEvent();
          Logger.log("Deleting event: " + event.getTitle())
        };
      }, 5, 500);
    });


    // Add newly fetched events if any
    try {
      var eventsToday = groupedEvents[d.toISOString().slice(0, 10)];
      Logger.log(eventsToday.length + " new events on " + d)

      for (var i = 0; i < eventsToday.length; i++) {
        var eventData = eventsToday[i];

        // Convert event
        var startTime = new Date(eventData.start);
        var endTime = new Date(eventData.end);

        var parts = eventData.title.split('#');

        // Check if the title has the expected number of parts
        if (parts.length >= 5) {
          // The first part after 'H'
          var eventName = parts[0].slice(1).trim();

          // Check against blacklist
          if (blacklist.includes(eventData.hash) || blacklist.includes(eventName)) {
            Logger.log("Skipping event: " + eventName);
            continue;
          }

          // The second part (class name)
          var className = parts[1].trim();
          // The third part (additional description)
          var additionalDescription = parts[2].slice(1).trim();
          var location = parts[4].substring(2, parts[4].length - 1).trim();

          var event = retryWithBackoff(function () {
            return targetCalendar.createEvent(eventName, startTime, endTime, {
              description: className + "\n" + eventData.description + "\n" + additionalDescription + "\nUnique Exclusion Hash: " + eventData.hash,
              location: location,
            });
          }, 5, 500); // 5 retries, starting with 500ms delay

          // Set tag with retries
          retryWithBackoff(function () {
            event.setTag("hka_scraper", "true");
          }, 5, 500);

          Logger.log("Added event: " + eventName)
        }
      }
    }
    catch (ex) {
      Logger.log("No new events on " + d)
    }
  }
}

// Function to retry with exponential backoff since google ratelimits us
function retryWithBackoff(fn, maxRetries, delay) {
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn(); // Try to execute the function
    } catch (e) {
      Logger.log("Error: " + e.toString());
      if (attempt === maxRetries) {
        throw new Error("Max retries reached: " + e.toString());
      }
      // Exponential backoff (2^attempt * delay ms)
      Utilities.sleep(Math.pow(2, attempt) * delay);
    }
  }
}

function uninstall(){
  deleteAllTriggers();
}

function deleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    if (["startSync","install","main","checkForUpdate"].includes(triggers[i].getHandlerFunction())){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function getValidTriggerFrequency(origFrequency) {
  if (origFrequency <= 0) {
    Logger.log("Specified frequency too fast. Defaulting to 15 minutes.");
    return 15;
  }
  
  var adjFrequency = Math.round(origFrequency/5) * 5; // Set the number to be the closest divisible-by-5
  adjFrequency = Math.max(adjFrequency, 1); // Make sure the number is at least 1 (0 is not valid for the trigger)
  adjFrequency = Math.min(adjFrequency, 15); // Make sure the number is at most 15 (will check for the 30 value below)
  
  if((adjFrequency == 15) && (Math.abs(origFrequency-30) < Math.abs(origFrequency-15)))
    adjFrequency = 30; // If we adjusted to 15, but the original number is actually closer to 30, set it to 30 instead
  
  if (origFrequency != adjFrequency) {
    Logger.log("Intended frequency = "+origFrequency+" adjusted to = "+adjFrequency);
  }

  return adjFrequency;
}