// Paste into script.google.com (Extensions > Apps Script in a new Google Sheet).
const ORGANIZERS = "theuntraditionallosers@gmail.com"; // get a note for every sign-up
const FILM = "Misbeehaved";
const WHEN = "Date and time TBD · FAU Theater #2"; // edit me
const TAB = "Reservations"; // columns: Time | Name | Email | Seat | Checked in

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(TAB);
  if (!s) { s = ss.insertSheet(TAB); s.appendRow(["Time","Name","Email","Seat","Checked in"]); }
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  return json_({ taken: rows.map(r => r[3]) });
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name||"").trim().slice(0,100), email = String(d.email||"").trim().toLowerCase().slice(0,200), seat = String(d.seat||"");
    if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !/^[A-Z]\d{1,2}$/.test(seat)) return json_({ok:false,error:"invalid"});
    const s = sheet_(), rows = s.getDataRange().getValues().slice(1);
    if (rows.some(r => r[3] === seat)) return json_({ok:false,error:"taken"});
    if (rows.some(r => String(r[2]).toLowerCase() === email)) return json_({ok:false,error:"duplicate"});
    s.appendRow([new Date(), name, email, seat, ""]);
    SpreadsheetApp.flush();
    MailApp.sendEmail(email, `You're confirmed: ${FILM}, seat ${seat}`,
      `Hi ${name},\n\nYour seat is reserved.\n\nFilm: ${FILM}\nWhen: ${WHEN}\nSeat: ${seat}\n\nShow this email at the door. If you can't make it, reply so we can free your seat.\n\nHosted by The Untraditional Losers\nFollow us: @theuntraditionallosers\n`);
    MailApp.sendEmail(ORGANIZERS, `New RSVP: ${name} (seat ${seat})`, `Name: ${name}\nEmail: ${email}\nSeat: ${seat}\n\nFull list is in the "${TAB}" sheet. Mark "Checked in" at the door.`);
    return json_({ok:true});
  } finally { lock.releaseLock(); }
}