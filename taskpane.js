

let meetingData = {
  subject: "Unavailable",
  body: "Unavailable",
  organizer: "Unavailable",
  meetingType: "",
  enableMom: "",
  startTime: "Unavailable",
  endTime: "Unavailable"
};

window.meetingData = meetingData; // ðŸ”¥ This is the fix


// Create a promise to track when Office is ready and initial values are fetched
const initReady = new Promise((resolve) => {
  Office.onReady(function (info) {
      console.log("ðŸ”„ Office.onReady info:", info);

      if (info.host === Office.HostType.Outlook) {
          console.log("âœ… Outlook Add-in initialized (Compose Mode)");

          const item = Office.context.mailbox.item;
          meetingData.organizer = Office.context.mailbox.userProfile.emailAddress || "Unavailable";
          console.log("ðŸ‘¤ Organizer:", meetingData.organizer);

          let initCount = 0;
          const totalInit = 4;

          const checkDone = () => {
              initCount++;
              if (initCount === totalInit) {
                  console.log("ðŸŸ¢ Initialization complete:", meetingData);
                  resolve();
              }
          };

          // Get subject
          item.subject.getAsync((res) => {
              if (res.status === Office.AsyncResultStatus.Succeeded) {
                  meetingData.subject = res.value || "Unavailable";
                  console.log("ðŸ“ Subject:", meetingData.subject);
              } else {
                  meetingData.subject = "Unavailable";
                  console.warn("âš ï¸ Failed to get subject:", res.error.message);
              }
              checkDone();
          });

          // Get body
          item.body.getAsync(Office.CoercionType.Text, (res) => {
              if (res.status === Office.AsyncResultStatus.Succeeded) {
                  meetingData.body = res.value || "Unavailable";
                  console.log("ðŸ§¾ Body:", meetingData.body);
              } else {
                  meetingData.body = "Unavailable";
                  console.warn("âš ï¸ Failed to get body:", res.error.message);
              }
              checkDone();
          });

          // Get start time
          if (item.start) {
              item.start.getAsync((res) => {
                  console.log("ðŸ“¥ Raw start time:", res.value);
                  if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
                      const localStart = new Date(res.value);
                      if (!isNaN(localStart)) {
                          meetingData.startTime = formatDateForMySQL(localStart);
                          console.log("ðŸ•’ Start Time:", meetingData.startTime);
                      } else {
                          console.warn("âŒ Invalid Start Date object:", res.value);
                          meetingData.startTime = "Unavailable";
                      }
                  } else {
                      console.warn("âš ï¸ Failed to get start time:", res.error?.message || "No value");
                      meetingData.startTime = "Unavailable";
                  }
                  checkDone();
              });
          } else {
              meetingData.startTime = "Unavailable";
              checkDone();
          }

          // Get end time
          if (item.end) {
              item.end.getAsync((res) => {
                  console.log("ðŸ“¥ Raw end time:", res.value);
                  if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
                      const localEnd = new Date(res.value);
                      if (!isNaN(localEnd)) {
                          meetingData.endTime = formatDateForMySQL(localEnd);
                          console.log("ðŸ•” End Time:", meetingData.endTime);
                      } else {
                          console.warn("âŒ Invalid End Date object:", res.value);
                          meetingData.endTime = "Unavailable";
                      }
                  } else {
                      console.warn("âš ï¸ Failed to get end time:", res.error?.message || "No value");
                      meetingData.endTime = "Unavailable";
                  }
                  checkDone();
              });
          } else {
              meetingData.endTime = "Unavailable";
              checkDone();
          }


          
          // Input listeners
const meetingTypeElem = document.getElementById("meetingType");
const momOptions = document.querySelectorAll('input[name="enableMom"]');

if (meetingTypeElem) {
  meetingTypeElem.addEventListener("change", () => {
      meetingData.meetingType = meetingTypeElem.value;
      console.log("ðŸ“Œ Meeting Type changed:", meetingData.meetingType);
  });
  meetingData.meetingType = meetingTypeElem.value;
  console.log("ðŸ“Œ Meeting Type default:", meetingData.meetingType);
}

const momChecked = document.querySelector('input[name="enableMom"]:checked');
meetingData.enableMom = momChecked ? momChecked.value : "";

momOptions.forEach((option) => {
  option.addEventListener("change", () => {
      if (option.checked) {
          meetingData.enableMom = option.value;
          console.log("ðŸ“Œ Enable MOM changed:", meetingData.enableMom);
      }
  });
});


          setInterval(updateMeetingData, 2000);
          sessionStorage.removeItem("hasSentData");
      }
  });
});

// Update meeting data periodically
function updateMeetingData() {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
          meetingData.subject = res.value || "Unavailable";
      }
  });

  item.body.getAsync(Office.CoercionType.Text, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
          meetingData.body = res.value || "Unavailable";
      }
  });

  if (item.start) {
      item.start.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
              const localStart = new Date(res.value);
              if (!isNaN(localStart)) {
                  meetingData.startTime = formatDateForMySQL(localStart);
              }
          }
      });
  }

  if (item.end) {
      item.end.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
              const localEnd = new Date(res.value);
              if (!isNaN(localEnd)) {
                  meetingData.endTime = formatDateForMySQL(localEnd);
              }
          }
      });
  }
}





// Debugging the sendBeacon call and async data fetching
window.addEventListener("unload", async () => {
  if (sessionStorage.getItem("hasSentData") === "true") return;
  sessionStorage.setItem("hasSentData", "true");

  try {
    if (!window.meetingData) {
      console.warn("â›” meetingData is undefined.");
      return;
    }

    const {
      meetingType,
      enableMom,
      subject,
      body,
      organizer,
      organizer_email,
      location,
      startTime,
      endTime,
      isOnlineMeeting,
      join_url,
      attendees
    } = meetingData;

    if (!meetingType || !enableMom) {
      console.warn("â›” Missing meetingType or enableMom, skipping auto-send.");
      return;
    }

    const requestData = {

      organizer: meetingData.organizer ,
      organizer_email: meetingData.organizer,
      subject: meetingData.subject || "",
      start: meetingData.startTime || "Unavailable",
      end: meetingData.endTime || "Unavailable",
      meeting_type: meetingData.meetingType || "Undefined",
      enable_mom: meetingData.enableMom || "Yes",
      preview:meetingData.body ||  "",
      location: location || "",
      isOnlineMeeting: isOnlineMeeting || false,
      join_url: join_url || "",
      attendees: attendees || []
    };

    console.log("ðŸ“¤ Preparing to send requestData:", requestData);
    
    // Send request via sendBeacon
    const beaconSent = navigator.sendBeacon(
      "https://add-in-gvbvabchhdf6h3ez.centralindia-01.azurewebsites.net/save-meeting/",
    // "https://addd-ycyt.onrender.com",
      new Blob([JSON.stringify(requestData)], { type: "application/json" })
    );
    
    if (beaconSent) {
      console.log("âœ… Data sent successfully via beacon.");
    } else {
      console.warn("âš ï¸ Failed to send data via beacon.");
    }
    
    // Update meeting body with data
    if (Office.context?.mailbox?.item) {
      const updateText = `Meeting Type: ${meetingType}\nEnable MOM: ${enableMom}`;
      Office.context.mailbox.item.body.setAsync(
        updateText,
        { coercionType: Office.CoercionType.Text },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("âœ… Meeting body updated on unload.");
          } else {
            console.warn("âŒ Failed to update body:", result.error.message);
          }
        }
      );
    }
  } catch (err) {
    console.error("âŒ Error in unload logic:", err);
  }
});

// âœ… Format date to ISO 8601 format (with timezone offset)
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