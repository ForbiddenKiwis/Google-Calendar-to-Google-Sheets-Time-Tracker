/**
 * Runs when the spreadsheet is opened and adds the menu options
 * to the spreadsheet menu
 */
const onOpen = () => {
    SpreadsheetApp.getUi().createMenu("Custom Menu")
        .addItem("Update Sheet Names", "setMenuName")
        .addToUi();
};

/**
 * Builds the cutomMenu obj and runs the sync
 */
function run() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Hours");
    const calendarId = sheet.getRange("B2").getValue(); 
    
    if (!calendarId) {
        SpreadsheetApp.getUi().alert("Please enter a Calendar ID in cell B2.");
        return;
    }

    syncCalendarEvents(calendarId);
}

function syncCalendarEvents(calendarId) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Hours");
    const calendar = CalendarApp.getCalendarById(calendarId);

    // Clear the prev data except the header row
    if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 4, sheet.getLastRow() - 1, 7).clearContent();
    }

    // Range of the calendar
    const allEvents = calendar.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1));

    if (allEvents.length === 0) {
        SpreadsheetApp.getUi().alert("No Events found in the calendar.");
        return;
    }

    const data = allEvents.map(event => [
        calendarId,
        event.getTitle(),
        event.getStartTime(),
        event.getEndTime(),
        '',
        '',
        calendar.getName()
    ]);

    // Write to sheet starting at D2
    sheet.getRange(2, 4, data.length, data[0].length).setValues(data);

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

    SpreadsheetApp.getUi().alert(`✅ Synced ${data.length} events from "${calendar.getName()}"`);
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

// /**
//  * The main function used for the sync 
//  * @param {Object} par the main parameter obj.
//  * @return {Object} The customMenu Object.
//  */
// const customMenu = (par) => {
//     const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//     const allHoursSheet = spreadsheet.getSheetByName("All Hours");


//     /**
//      * Fomat the sheet
//      */    
//     const formatSheet = () => {
//         allHoursSheet.sort(5, false);

//         if (allHoursSheet.getLastRow() > 1 && allHoursSheet.getLastRow() < allHoursSheet.getMaxRows()) {
//             allHoursSheet.deleteRows(allHoursSheet.getLastRow() + 1, allHoursSheet.getMaxRows() - allHoursSheet.getLastRow());
//         }

//         // Set the Time Duration Formula in G cell. 
//         // Formula: End (F) - Start (E), showing time duration in HH:MM
//         allHoursSheet.getRange('G2:G').setFormulaR1C1(
//             '=IF(OR(R[0]C[-2]="", R[0]C[-1]=""), "", R[0]C[-1] - R[0]C[-2])'
//         );

//         // Set Numeric Duration formula in cell H. 
//         // Formula: Time Duration (G) * 24 to convert time to decimal hours
//         allHoursSheet.getRange('H2:H').setFormulaR1C1(
//             '=IF(R[0]C[-1]="", "", R[0]C[-1] * 24)'
//         );
//     };

//     function onEdit(e) {
//         allHoursSheet.sort(5, false);

//         if (allHoursSheet.getLastRow() > 1 &&
//             allHoursSheet.getLastRow() < allHoursSheet.getMaxRows()) {
//         allHoursSheet.deleteRows(allHoursSheet.getLastRow() + 1, allHoursSheet.getMaxRows() - allHoursSheet.getLastRow());
//         }

//         allHoursSheet.getRange('G2:G').setFormulaR1C1(
//         '=IF(OR(R[0]C[-2]="", R[0]C[-1]=""), "", R[0]C[-1] - R[0]C[-2])'
//         );

//         allHoursSheet.getRange('H2:H').setFormulaR1C1(
//         '=IF(R[0]C[-1]="", "", R[0]C[-1] * 24)'
//         );
//     }

//     /**
//      * Sync calendar data to "All Hours" sheet
//      */

//     const syncSingleCalendar = (calendarId) => {
//         const calendar = CalendarApp.getCalendarById(calendarId);
//         if (!calendar) {
//             Logger.log('Calendar not found: ' + calendarId);
//             return;
//         }

//         const calendarName = calendar.getName();
//         const allEvents = calendar.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1));

//         if (allEvents.length === 0) {
//             Logger.log('No events found for calendar: ' + calendarId);
//             return;
//         }

//         const syncStartDate = new Date(Math.min(...allEvents.map(e => e.getStartTime().getTime())));
//         const syncEndDate = new Date(Math.max(...allEvents.map(e => e.getEndTime().getTime())));
//         const events = calendar.getEvents(syncStartDate, syncEndDate);

//         const existingEvents = allHoursSheet.getDataRange().getValues().slice(1);
//         const existingEventsLookup = existingEvents.reduce((lookup, row, index) => {
//             if (row[0] === calendarId) {
//                 lookup[row[1]] = { row: index + 2, data: row };
//             }
//             return lookup;
//         }, {});

//         events.forEach(event => {
//             const eventId = event.getId();
//             const existing = existingEventsLookup[eventId];

//             const rowData = [
//                 calendarId,
//             ]
//         });
//     }
// }

