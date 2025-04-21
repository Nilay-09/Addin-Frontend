// Initialize meeting data with default values
let meetingData = {
    subject: "Unavailable",
    body: "Unavailable",
    organizer: "Unavailable",
    meetingType: "Standup", // Default value
    enableMom: "Yes",      // Default value
    startTime: "Unavailable",
    endTime: "Unavailable",
    location: "",
    isOnlineMeeting: false,
    join_url: "",
    attendees: []
  };
  
  // Make meetingData available in global scope
  window.meetingData = meetingData;
  
  // This flag ensures we only initialize once per page load
  let isInitialized = false;
  let dataLoadedPromise = null;
  
  // Create a promise to track when Office is ready and initial values are fetched
  Office.onReady(function(info) {
    console.log("📝 Office.onReady info:", info);
    
    if (info.host === Office.HostType.Outlook) {
      console.log("✅ Outlook Add-in initialized");
      
      if (!isInitialized) {
        initializeAddin();
      }
    }
  });
  
  // Initialize the add-in - called once when Office is ready
  // force parameter allows re-initialization when needed
  window.initializeAddin = function(force = false) {
    if (isInitialized && !force) {
      console.log("📝 Add-in already initialized, skipping");
      return;
    }
    
    console.log("📝 Initializing add-in" + (force ? " (forced)" : ""));
    isInitialized = true;
    
    const item = Office.context.mailbox.item;
    meetingData.organizer = Office.context.mailbox.userProfile.emailAddress || "Unavailable";
    console.log("👤 Organizer:", meetingData.organizer);
    
    // Create a promise to track when data is loaded
    dataLoadedPromise = new Promise((resolve) => {
      // Load any previously saved custom properties
      loadSavedProperties().then(() => {
        // Get real-time data
        return loadItemData();
      }).then(() => {
        resolve();
      });
    });
    
    // Set up UI listeners for the form if we're in the taskpane
    setupUIListeners();
    
    // Reset sessionStorage flag for sending data
    sessionStorage.removeItem("hasSentData");
    
    // Handle window unload (when taskpane closes)
    window.addEventListener("unload", handleUnload);
  };
  
  // Load any previously saved custom properties
  function loadSavedProperties() {
    return new Promise((resolve) => {
      const item = Office.context.mailbox.item;
      
      item.loadCustomPropertiesAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("⚠️ Failed to load custom properties:", result.error.message);
          resolve();
          return;
        }
        
        const props = result.value;
        const savedData = props.get("meetingFormData");
        
        if (savedData) {
          try {
            const parsedData = JSON.parse(savedData);
            console.log("📤 Loaded saved properties:", parsedData);
            
            // Update meetingData with saved values
            meetingData.meetingType = parsedData.meetingType || "Standup";
            meetingData.enableMom = parsedData.enableMom || "Yes";
            
            // Update UI if possible
            updateUIFromData();
            resolve();
          } catch (e) {
            console.error("❌ Error parsing saved properties:", e);
            resolve();
          }
        } else {
          console.log("ℹ️ No saved properties found, using defaults");
          resolve();
        }
      });
    });
  }
  
  // Save meeting form selections to custom properties
  function saveFormData() {
    const item = Office.context.mailbox.item;
    
    item.loadCustomPropertiesAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("⚠️ Failed to load custom properties for saving:", result.error.message);
        return;
      }
      
      const props = result.value;
      const dataToSave = {
        meetingType: meetingData.meetingType,
        enableMom: meetingData.enableMom
      };
      
      props.set("meetingFormData", JSON.stringify(dataToSave));
      
      props.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ Form data saved to custom properties");
        } else {
          console.warn("⚠️ Failed to save form data:", saveResult.error.message);
        }
      });
    });
  }
  
  // Set up UI event listeners if the form elements exist
  function setupUIListeners() {
    const meetingTypeElem = document.getElementById("meetingType");
    const momOptions = document.querySelectorAll('input[name="enableMom"]');
    
    // Set up meeting type dropdown listener
    if (meetingTypeElem) {
      // First set the UI to match our data
      meetingTypeElem.value = meetingData.meetingType;
      
      // Then add the change listener
      meetingTypeElem.addEventListener("change", () => {
        meetingData.meetingType = meetingTypeElem.value;
        console.log("🔄 Meeting Type changed:", meetingData.meetingType);
        saveFormData();
      });
    }
    
    // Set up MOM radio button listeners
    if (momOptions.length > 0) {
      // First set the UI to match our data
      momOptions.forEach((option) => {
        if (option.value === meetingData.enableMom) {
          option.checked = true;
        }
      });
      
      // Then add change listeners
      momOptions.forEach((option) => {
        option.addEventListener("change", () => {
          if (option.checked) {
            meetingData.enableMom = option.value;
            console.log("🔄 Enable MOM changed:", meetingData.enableMom);
            saveFormData();
          }
        });
      });
    }
  }
  
  // Update UI elements with current data values (if they exist)
  function updateUIFromData() {
    const meetingTypeElem = document.getElementById("meetingType");
    const momOptions = document.querySelectorAll('input[name="enableMom"]');
    
    if (meetingTypeElem) {
      meetingTypeElem.value = meetingData.meetingType;
    }
    
    if (momOptions.length > 0) {
      momOptions.forEach((option) => {
        if (option.value === meetingData.enableMom) {
          option.checked = true;
        }
      });
    }
  }
  
  // Load item data from the current Outlook item
  function loadItemData() {
    return new Promise((resolve) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        resolve();
        return;
      }
      
      let pendingOperations = 0;
      const trackOperation = () => {
        pendingOperations++;
        return () => {
          pendingOperations--;
          if (pendingOperations === 0) {
            resolve();
          }
        };
      };
      
      // Get subject
      trackOperation();
      item.subject.getAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          meetingData.subject = res.value || "Unavailable";
        }
        trackOperation()();
      });
      
      // Get body
      trackOperation();
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          meetingData.body = res.value || "Unavailable";
        }
        trackOperation()();
      });
      
      // Get start time
      if (item.start) {
        trackOperation();
        item.start.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
            const localStart = new Date(res.value);
            if (!isNaN(localStart)) {
              meetingData.startTime = formatDateForMySQL(localStart);
            }
          }
          trackOperation()();
        });
      }
      
      // Get end time
      if (item.end) {
        trackOperation();
        item.end.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
            const localEnd = new Date(res.value);
            if (!isNaN(localEnd)) {
              meetingData.endTime = formatDateForMySQL(localEnd);
            }
          }
          trackOperation()();
        });
      }
      
      // Get location
      if (item.location) {
        trackOperation();
        item.location.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            meetingData.location = res.value || "";
          }
          trackOperation()();
        });
      }
      
      // For online meetings, try to get the join URL
      if (typeof item.isOnlineMeeting !== 'undefined') {
        meetingData.isOnlineMeeting = item.isOnlineMeeting;
        
        if (item.meetingUrl) {
          meetingData.join_url = item.meetingUrl || "";
        }
      }
      
      // Try to get attendees if available
      if (item.requiredAttendees) {
        trackOperation();
        item.requiredAttendees.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
            meetingData.attendees = res.value.map(attendee => attendee.emailAddress) || [];
          }
          trackOperation()();
        });
      }
      
      // If there were no async operations started, resolve immediately
      if (pendingOperations === 0) {
        resolve();
      }
    });
  }
  
  // Handle window unload event (when taskpane closes)
  function handleUnload() {
    console.log("🚪 Taskpane unloading, checking if we need to send data");
    sendMeetingData();
  }
  
  // Send meeting data to the API
  function sendMeetingData(forceRefresh = false) {
    // If we're forcing a refresh, don't check the sessionStorage flag
    if (!forceRefresh && sessionStorage.getItem("hasSentData") === "true") {
      console.log("🔄 Data was already sent this session, skipping");
      return Promise.resolve();
    }
    
    sessionStorage.setItem("hasSentData", "true");
    
    // If we need to force refresh data first
    const sendData = () => {
      try {
        const item = Office.context.mailbox.item;
        
        // Always use the latest data
        const requestData = {
          organizer: meetingData.organizer,
          organizer_email: meetingData.organizer,
          subject: meetingData.subject || "",
          start: meetingData.startTime || "Unavailable",
          end: meetingData.endTime || "Unavailable",
          meeting_type: meetingData.meetingType || "Standup",
          enable_mom: meetingData.enableMom || "Yes",
          preview: meetingData.body || "",
          location: meetingData.location || "",
          isOnlineMeeting: meetingData.isOnlineMeeting || false,
          join_url: meetingData.join_url || "",
          attendees: meetingData.attendees || []
        };
        
        console.log("📤 Sending meeting data:", requestData);
        
        // Use sendBeacon for more reliable delivery during page unload
        const beaconSent = navigator.sendBeacon(
          "https://add-in-gvbvabchhdf6h3ez.centralindia-01.azurewebsites.net/save-meeting/",
          new Blob([JSON.stringify(requestData)], { type: "application/json" })
        );
        
        if (beaconSent) {
          console.log("✅ Data sent successfully via beacon");
          return Promise.resolve();
        } else {
          console.warn("⚠️ Failed to send data via beacon");
          
          // Fallback to fetch API if sendBeacon fails
          return fetch("https://add-in-gvbvabchhdf6h3ez.centralindia-01.azurewebsites.net/save-meeting/", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify(requestData),
            keepalive: true
          })
          .then(response => {
            console.log("✅ Data sent successfully via fetch", response.status);
          })
          .catch(error => {
            console.error("❌ Error sending data via fetch:", error);
          });
        }
      } catch (err) {
        console.error("❌ Error in sendMeetingData:", err);
        return Promise.reject(err);
      }
    };
    
    if (forceRefresh) {
      // Make sure we load the latest data before sending
      return loadItemData().then(sendData);
    } else {
      // Use existing data
      return sendData();
    }
  }
  
  // Format date to ISO 8601 format (with timezone offset)
  function formatDateForMySQL(date) {
    const pad = (n) => (n < 10 ? '0' + n : n);
    
    // Get timezone offset in minutes and convert it to hours and minutes
    const timezoneOffset = date.getTimezoneOffset();
    const offsetHours = String(Math.floor(Math.abs(timezoneOffset) / 60)).padStart(2, '0');
    const offsetMinutes = String(Math.abs(timezoneOffset) % 60).padStart(2, '0');
    const offsetSign = timezoneOffset > 0 ? '-' : '+';
    
    // Format date as "YYYY-MM-DDTHH:mm:ss+/-HH:mm"
    return (
      date.getFullYear() + '-' +
      pad(date.getMonth() + 1) + '-' +
      pad(date.getDate()) + 'T' +
      pad(date.getHours()) + ':' +
      pad(date.getMinutes()) + ':' +
      pad(date.getSeconds()) +
      offsetSign + offsetHours + ':' + offsetMinutes
    );
  }
  
  // Expose function for event handlers in functions.html to access
  window.getMeetingDataAndSend = function(forceRefresh = false) {
    // Make sure we're initialized
    if (!isInitialized) {
      window.initializeAddin(true);
    }
    
    // Set a small timeout to ensure we've loaded the latest data
    if (forceRefresh) {
      return loadItemData().then(() => {
        return sendMeetingData(true);
      });
    } else {
      if (dataLoadedPromise) {
        return dataLoadedPromise.then(() => {
          return sendMeetingData();
        });
      } else {
        return sendMeetingData();
      }
    }
  };