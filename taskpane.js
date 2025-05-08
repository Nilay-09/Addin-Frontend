

// taskpane.js - Merged implementation combining correct state management and reliable API calls

// Initialize meeting data with default values
let meetingData = {
  subject: "Unavailable",
  body: "Unavailable",
  organizer: "Unavailable",
  meetingType: "Standup",
  enableMom: "Yes",
  startTime: "Unavailable",
  endTime: "Unavailable",
  location: "",
  isOnlineMeeting: false,
  join_url: "",
  attendees: []
};

// Expose meetingData globally
window.meetingData = meetingData;
let isInitialized = false;
let dataLoadedPromise = null;

// Office add-in entry point
Office.onReady(info => {
  console.log("📝 Office.onReady info:", info);
  if (info.host === Office.HostType.Outlook) {
    initializeAddin();
  }
});

// Initialize or re-initialize the add-in
function initializeAddin(force = false) {
  if (isInitialized && !force) {
    console.log("📝 Already initialized, skipping");
    return;
  }
  console.log(`📝 Initializing add-in${force ? ' (forced)' : ''}`);
  isInitialized = true;
  // Reset send flag each session
  sessionStorage.removeItem("hasSentData");

  // Set organizer
  meetingData.organizer = Office.context.mailbox.userProfile.emailAddress || "Unavailable";

  // Load saved properties then live item data
  dataLoadedPromise = loadSavedProperties()
    .then(loadItemData)
    .then(updateUIFromData);

  // Setup UI listeners
  setupUIListeners();
}

// Load persisted form selections
function loadSavedProperties() {
  return new Promise(resolve => {
    Office.context.mailbox.item.loadCustomPropertiesAsync(res => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const saved = res.value.get("meetingFormData");
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            meetingData.meetingType = parsed.meetingType || meetingData.meetingType;
            meetingData.enableMom = parsed.enableMom || meetingData.enableMom;
            console.log("📤 Loaded saved selections:", parsed);
          } catch (e) {
            console.error("Error parsing savedProperties", e);
          }
        }
      }
      resolve();
    });
  });
}

// Persist current form selections
function saveFormData() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(res => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      const props = res.value;
      props.set("meetingFormData", JSON.stringify({
        meetingType: meetingData.meetingType,
        enableMom: meetingData.enableMom
      }));
      props.saveAsync(saveRes => console.log(
        saveRes.status === Office.AsyncResultStatus.Succeeded ?
          "✅ Form data saved" :
          `⚠️ Failed to save (${saveRes.error.message})`
      ));
    }
  });
}

// Attach UI listeners once
function setupUIListeners() {
  const mt = document.getElementById("meetingType");
  if (mt && !mt._listener) {
    mt._listener = true;
    // Initialize dropdown value
    mt.value = meetingData.meetingType;
    mt.addEventListener("change", () => {
      meetingData.meetingType = mt.value;
      console.log("🔄 Meeting Type:", mt.value);
      saveFormData();
    });
  }
  document.querySelectorAll('input[name="enableMom"]').forEach(opt => {
    if (!opt._listener) {
      opt._listener = true;
      if (opt.value === meetingData.enableMom) opt.checked = true;
      opt.addEventListener("change", () => {
        if (opt.checked) {
          meetingData.enableMom = opt.value;
          console.log("🔄 Enable MOM:", opt.value);
          saveFormData();
        }
      });
    }
  });
}

// Update form inputs from meetingData
function updateUIFromData() {
  const mt = document.getElementById("meetingType");
  if (mt) {
    const optionExists = Array.from(mt.options).some(o => o.value === meetingData.meetingType);
    mt.value = optionExists ? meetingData.meetingType : mt.options[0].value;
    if (!optionExists) {
      meetingData.meetingType = mt.value;
      saveFormData();
    }
  }
  document.querySelectorAll('input[name="enableMom"]').forEach(opt => {
    opt.checked = (opt.value === meetingData.enableMom);
  });
}

// Load live item details
function loadItemData() {
  return new Promise(resolve => {
    const item = Office.context.mailbox.item;
    if (!item) return resolve();
    let pending = 0;
    const done = () => (--pending === 0) && resolve();
    const track = () => { pending++; return done; };

    track(); item.subject.getAsync(r => { meetingData.subject = r.value || meetingData.subject; done(); });
    track(); item.body.getAsync(Office.CoercionType.Text, r => { meetingData.body = r.value || meetingData.body; done(); });
    if (item.start) { track(); item.start.getAsync(r => { meetingData.startTime = formatDate(r.value); done(); }); }
    if (item.end)   { track(); item.end.getAsync(r => { meetingData.endTime = formatDate(r.value); done(); }); }
    if (item.location) { track(); item.location.getAsync(r => { meetingData.location = r.value; done(); }); }
    meetingData.isOnlineMeeting = item.isOnlineMeeting || false;
    meetingData.join_url = item.meetingUrl || meetingData.join_url;
    if (item.requiredAttendees) {
      track(); item.requiredAttendees.getAsync(r => {
        meetingData.attendees = r.value.map(a => a.emailAddress);
        done();
      });
    }
    if (pending === 0) resolve();
  });
}

// Send meeting data to the API, with beacon + fetch fallback and session control
function sendMeetingData(forceRefresh = false) {
  if (!forceRefresh && sessionStorage.getItem("hasSentData") === "true") {
    console.log("🔄 Data already sent this session, skipping");
    return Promise.resolve();
  }
  sessionStorage.setItem("hasSentData", "true");

//   const url = "https://dfmmt-b0enaqg6eyg9ewfv.centralindia-01.azurewebsites.net/save-meeting/";
  const url = "https://add-in-gvbvabchhdf6h3ez.centralindia-01.azurewebsites.net/save-meeting/";
  const payload = {
    organizer:       meetingData.organizer,
    organizer_email: meetingData.organizer,
    subject:         meetingData.subject,
    start:           meetingData.startTime,
    end:             meetingData.endTime,
    meeting_type:    meetingData.meetingType,
    enable_mom:      meetingData.enableMom,
    preview:         meetingData.body,
    location:        meetingData.location,
    isOnlineMeeting: meetingData.isOnlineMeeting,
    join_url:        meetingData.join_url,
    attendees:       meetingData.attendees
  };

  const blob = new Blob([JSON.stringify(payload)], { type: "application/json" });
  const doSend = () => {
    const beaconOk = navigator.sendBeacon(url, blob);
    if (beaconOk) {
      console.log("✅ Data queued via sendBeacon");
      console.log("📤 Payload :", payload);
      return Promise.resolve();
    }
    console.warn("⚠️ sendBeacon failed, using fetch fallback");
    return fetch(url, {
      method:      'POST',
      mode:        'cors',
      credentials: 'omit',
      headers:     { 'Content-Type': 'application/json' },
      body:        JSON.stringify(payload),
      keepalive:   true
    })
    .then(res => console.log("✅ Fallback fetch status", res.status))
    .catch(err => console.error("❌ Fetch fallback error", err));
  };

  if (forceRefresh) {
    return loadItemData().then(doSend);
  }
  return doSend();
}

// Trigger save on pane close
window.addEventListener('beforeunload', () => sendMeetingData(true));
window.addEventListener('unload',        () => sendMeetingData(true));

// Expose for event-based triggers
window.getMeetingDataAndSend = function(force = false) {
  if (!isInitialized || force) initializeAddin(force);
  const ready = dataLoadedPromise || Promise.resolve();
  return ready.then(() => sendMeetingData(true));
};

// Helper to format dates to ISO 8601 with timezone
function formatDate(raw) {
  const d = new Date(raw);
  const pad = n => n < 10 ? '0'+n : n;
  const tz = d.getTimezoneOffset();
  const sign = tz > 0 ? '-' : '+';
  const hh = String(Math.floor(Math.abs(tz)/60)).padStart(2,'0');
  const mm = String(Math.abs(tz)%60).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}` +
         `T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}` +
         `${sign}${hh}:${mm}`;
}

// Make initializeAddin globally accessible
window.initializeAddin = initializeAddin;
