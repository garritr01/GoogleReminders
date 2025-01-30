function onOpen() {
	var ui = SpreadsheetApp.getUi();
	ui.createMenu('Custom Tools')
		.addItem('Move Complete Tasks', 'complete')
		.addItem('Order Rows by Due Date', 'orderByDate')
		.addItem('Expand Rows to Fit Text', 'expandRows')
		.addItem('Shrink Rows', 'shrinkRows')
		.addToUi();
}

function complete() {

	// Get currently opened sheet to display information about script running
	const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
	currentSheet.toast("Completing tasks...")

	// Get "To Do" and "Done" sheets, their headers, and data
	const srcSheetName = "To Do";
	const srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(srcSheetName);
	const srcValues = srcSheet.getDataRange().getValues();
	const srcHeader = srcValues[0];
	const srcData = srcValues.slice(1);

	const destSheetName = "Done";
	const destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(destSheetName);

	// Get a timestamp to record completion time
	const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd-yyyy-HH-mm');

	// Get rows and indices with status 1
	const statusIndex = srcHeader.indexOf("Status");
	const rowsToMove = [];
	const movedIndices = [];
	for (let i = 1; i < srcData.length; i++) {
		if (srcData[i][statusIndex] == 1) {
			// Replace completion status with completion timestamp
			const stampedRow = [...srcData[i].slice(0, srcData[i].length - 1), timestamp];
			rowsToMove.push(stampedRow);
			// Row index is srcData index + 1
			movedIndices.push(i);
		}
	}

	// Insert rows to top of "Done"
	for (row of rowsToMove) {
		destSheet.insertRowBefore(2);
		destSheet.getRange(2, 1, 1, row.length).setValues([row]);
	}

	// Delete rows from "To-Do"
	for (index of movedIndices.reverse()) {
		// +2 due to 1 indexed sheet and removed header
		srcSheet.deleteRow(index + 2);
	}

	currentSheet.toast(`Rows inserted into "Done" and removed from "To Do": ${movedIndices.map(i => i + 2).join(', ')}`);
}

function orderByDate() {
	const srcSheetName = "To Do"
	var srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(srcSheetName);

	SpreadsheetApp.getActiveSpreadsheet().toast("Ordering rows by due date...")

	var range = srcSheet.getDataRange();
	var data = range.getValues();
	var header = data[0];
	var rows = data.slice(1);

	const dueIndex = header.indexOf("Due Date");
	rows.sort(function (a, b) {
		var dateA = parseDate(a[dueIndex]);
		var dateB = parseDate(b[dueIndex]);
		return dateA - dateB;
	});

	var sortedData = [header].concat(rows);
	range.setValues(sortedData);
	SpreadsheetApp.getActiveSpreadsheet().toast("Rows ordered by due date.")
}

function parseDate(dateStr) {
	var parts = dateStr.split('-');

	if (parts.length === 5) {
		return new Date(parts[2], parts[0] - 1, parts[1], parts[3], parts[4]);
	} else {
		return new Date(0); // Return an invalid date if the format is not correct
	}
}

function expandRows() {
	const sheet = SpreadsheetApp.getActiveSheet();
	sheet.getRange(1, 2, sheet.getLastRow(), sheet.getLastColumn()).setWrap(true);
	sheet.autoResizeRows(1, sheet.getLastRow());
}

function shrinkRows() {
	const sheet = SpreadsheetApp.getActiveSheet();
	sheet.getRange(1, 2, sheet.getLastRow(), sheet.getLastColumn()).setWrap(false);
	sheet.autoResizeRows(1, sheet.getLastRow());
}