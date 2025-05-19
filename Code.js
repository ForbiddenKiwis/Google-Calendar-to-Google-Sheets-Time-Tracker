/**
 * Runs when the spreadsheet is opened and adds the menu options
 * to the spreadsheet menu
 */
const onOpen = () => {
    SpreadsheetApp.getUi().createMenu("Custom Menu")
        .addItem("Update Sheet Names", "setMenuName")
        .addItem("Sync iCal Events", "run")
        .addToUi();
};

/**
 * Builds the cutomMenu obj and runs the sync
 */
function run() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Hours");
    const icalUrl = sheet.getRange("B2").getValue(); 
    
    if (!icalUrl || !icalUrl.startsWith("http")) {
        SpreadsheetApp.getUi().alert("Please enter a valid ical URL in cell B2.");
        return;
    }

    try {
        syncICalEvents(icalUrl);
    } catch (error) {
        SpreadsheetApp.getUi().alert("❌ Error: " + error.message);
    }
}

function syncICalEvents(icalUrl) {
    const response = UrlFetchApp.fetch(icalUrl);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Hours");

    // Clear the prev data except the header row
    if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 4, sheet.getLastRow() - 1, 7).clearContent();
    }

    const content = response.getContentText();

    const events = [];
    const lines = content.split("\n");

    let currentEvent = {};
    
    for (let line of lines) {
        line = line.trim();

        if (line === "BEGIN:VEVENT") {
            currentEvent = {};
        } else if (line === "END:VEVENT") {
            if (currentEvent.summary && currentEvent.start && currentEvent.end) {
                events.push([
                currentEvent.summary,
                new Date(currentEvent.start),
                new Date(currentEvent.end),
                '',
                '',
                '',
                ]);
            }
        } else if (line.startsWith("SUMMARY:")) {
            currentEvent.summary = line.substring(8);
        } else if (line.startsWith("DTSTART")) {
            currentEvent.start = extractICalDate(line);
        } else if (line.startsWith("DTEND")) {
            currentEvent.end = extractICalDate(line);
        }
    }

    if (events.length === 0) {
        SpreadsheetApp.getUi().alert("No events found in the iCal feed.");
        return;
    }

    // Write to sheet starting at D2
    sheet.getRange(2, 4, events.length, events[0].length).setValues(events);

    /**
     * Set duration Formulas
     */ 

    //Time Duration
    const lastRow = sheet.getLastRow();
    sheet.getRange(2, 7, lastRow - 1).setFormulaR1C1(
        '=IF(OR(RC[-1]="", RC[-2]=""), "", RC[-1] - RC[-2])'
    );

    // Numeric Duration
    sheet.getRange(2,8, lastRow - 1).setFormulaR1C1(
        '=IF(RC[-1]="", "", RC[-1] * 24)'
    );

    SpreadsheetApp.getUi().alert(`✅ Synced ${events.length} events from events from iCal.`);
}

function setMenuName() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const newName = Browser.inputBox("Enter a new name for this sheet:")

    if (newName && newName.trim() !== "") {
        sheet.setName(newName);
        SpreadsheetApp.getUi().alert(`Sheet renamed to "${newName}"`);
    } else {
        SpreadsheetApp.getUi().alert("Invalid name. Sheet not renamed.");
    }
}

function extractICalDate(line) {
    const parts = line.split(":");
    if (parts.length !== 2) return null;

    const raw = parts[1].trim();

    // Support for both all-day (YYYYMMDD) and datetime (YYYYMMDDTHHMMSSZ) formats
    if (raw.length === 8) {
        return new Date(`${raw.substring(0,4)}-${raw.substring(4,6)}-${raw.substring(6,8)}`);
    } else if (raw.length >= 15) {
        return new Date(
            `${raw.substring(0,4)}-${raw.substring(4,6)}-${raw.substring(6,8)}T${raw.substring(9,11)}:${raw.substring(11,13)}:${raw.substring(13,15)}Z`
        );
    }

    return null;
}