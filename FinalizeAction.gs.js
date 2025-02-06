function finalizeAction() {
  var lendingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Lending'); // Lending sheet
  var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory'); // Inventory sheet
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History'); // History sheet
  var onHoldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('On Hold'); // On Hold sheet
  var ui = SpreadsheetApp.getUi(); // UI for displaying error messages
  var userEmail = Session.getActiveUser().getEmail(); // Get the email of the user who processed the action (Processed by)

  // Loop through all rows in the Lending sheet (assuming the form is starting from row 2)
  var lastRow = lendingSheet.getLastRow();

  for (var row = 2; row <= lastRow; row++) {
    var action = lendingSheet.getRange(row, 6).getValue(); // Action column in Lending sheet (Column F)
    var itemID = lendingSheet.getRange(row, 1).getValue(); // Item ID column (Column A)
    var itemName = lendingSheet.getRange(row, 2).getValue(); // Item Name column (Column B)
    var borrowerEmail = lendingSheet.getRange(row, 3).getValue(); // Borrower Email column (Column C)
    var borrowerName = lendingSheet.getRange(row, 4).getValue(); // Borrower Name column (Column D)
    var borrowerClass = lendingSheet.getRange(row, 5).getValue(); // Borrower Class column (Column E)

    // Validate that Borrower Email and Borrower Name are provided for Lend and Hold actions
    if ((action === 'Lend Item' || action === 'Hold Item') && (!borrowerEmail || !borrowerName)) {
      ui.alert('Error: Please provide both the borrower\'s email and name before proceeding.');
      return; // Exit the function to prevent further processing
    }

    // Find the corresponding item in the Inventory sheet
    var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow(), 12).getValues(); // Full Inventory sheet range with added Student Name and Class columns
    var onHoldColumnIndex = 11; // Assuming "On Hold?" is in column 12 (index 11)

    for (var i = 0; i < inventoryData.length; i++) {
      if (inventoryData[i][0] == itemID) { // Match the Item ID in Inventory (Column A)

        var currentStatus = inventoryData[i][4]; // Status column in Inventory (Column E)

        // Action: Lend Item
        if (action === 'Lend Item') {
          if (inventoryData[i][onHoldColumnIndex].toString().trim() === 'Yes') {
            ui.alert('Item is on Hold. Review "On Hold" and remove the hold before lending.');
            return; // Exit if the item is on hold
          }

          if (currentStatus === 'On Lend') {
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is already on loan.');
            return;
          }

          // Clear item from On Hold sheet if previously on hold
          var onHoldData = onHoldSheet.getRange(2, 1, onHoldSheet.getLastRow(), 6).getValues();
          for (var j = 0; j < onHoldData.length; j++) {
            if (onHoldData[j][0] == itemID) {
              onHoldSheet.deleteRow(j + 2);
              break;
            }
          }

          inventorySheet.getRange(i + 2, 5).setValue('On Lend'); // Update Status to "On Lend"
          inventorySheet.getRange(i + 2, 7).setValue(borrowerEmail); // Update Current Borrower
          inventorySheet.getRange(i + 2, 8).setValue(borrowerName); // Update Student Name
          inventorySheet.getRange(i + 2, 9).setValue(borrowerClass); // Update Class
          inventorySheet.getRange(i + 2, 10).setValue(new Date()); // Update Loan Date
          inventorySheet.getRange(i + 2, 11).setValue(new Date(new Date().setDate(new Date().getDate() + 7))); // Expected Return Date

          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Lend Item', new Date(), '', userEmail);

        // Action: Return Item
        } else if (action === 'Return Item') {
          if (currentStatus === 'In Stock') {
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is already in stock.');
            return;
          }

          inventorySheet.getRange(i + 2, 5).setValue('In Stock'); // Update Status to "In Stock"
          inventorySheet.getRange(i + 2, 7).setValue(''); // Clear Current Borrower
          inventorySheet.getRange(i + 2, 8).setValue(''); // Clear Student Name
          inventorySheet.getRange(i + 2, 9).setValue(''); // Clear Class
          inventorySheet.getRange(i + 2, 10).setValue(''); // Clear Loan Date
          inventorySheet.getRange(i + 2, 11).setValue(''); // Clear Expected Return Date

          var returnDate = new Date();
          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Return Item', '', returnDate, userEmail);

        // Action: Hold Item
        } else if (action === 'Hold Item') {
          var onHoldData = onHoldSheet.getRange(2, 1, onHoldSheet.getLastRow(), 6).getValues();
          var isAlreadyOnHold = onHoldData.some(entry => entry[0] == itemID);

          if (isAlreadyOnHold) {
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is already on hold.');
            return;
          }

          inventorySheet.getRange(i + 2, onHoldColumnIndex + 1).setValue('Yes'); // Set "On Hold?" to Yes

          var onHoldLastRow = onHoldSheet.getLastRow() + 1;
          onHoldSheet.getRange(onHoldLastRow, 1).setValue(itemID);
          onHoldSheet.getRange(onHoldLastRow, 2).setValue(itemName);
          onHoldSheet.getRange(onHoldLastRow, 3).setValue(borrowerEmail);
          onHoldSheet.getRange(onHoldLastRow, 4).setValue(borrowerName);
          onHoldSheet.getRange(onHoldLastRow, 5).setValue(borrowerClass);
          onHoldSheet.getRange(onHoldLastRow, 6).setValue(new Date()); // Correct Date Requested placement

          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Hold Item', '', '', userEmail);

        // Action: Remove Hold
        } else if (action === 'Remove Hold') {
          if (inventoryData[i][onHoldColumnIndex].toString().trim() !== 'Yes') {
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is not currently on hold.');
            return;
          }

          inventorySheet.getRange(i + 2, onHoldColumnIndex + 1).setValue('No'); // Remove Hold

          // Find and remove from On Hold sheet
          var onHoldData = onHoldSheet.getRange(2, 1, onHoldSheet.getLastRow(), 6).getValues();
          for (var j = 0; j < onHoldData.length; j++) {
            if (onHoldData[j][0] == itemID) {
              onHoldSheet.deleteRow(j + 2); // Remove the hold entry
              break;
            }
          }

          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Remove Hold', '', '', userEmail);
        }

        // Clear Lending Sheet fields
        lendingSheet.getRange(row, 6).setValue('');
        lendingSheet.getRange(row, 3).setValue('');
        lendingSheet.getRange(row, 4).setValue('');
        lendingSheet.getRange(row, 5).setValue('');

        break;
      }
    }
  }
}

function logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, action, dueDate, returnDate, userEmail) {
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  var historyLastRow = historySheet.getLastRow() + 1;

  historySheet.getRange(historyLastRow, 1).setValue(new Date());
  historySheet.getRange(historyLastRow, 2).setValue(itemID);
  historySheet.getRange(historyLastRow, 3).setValue(itemName);
  historySheet.getRange(historyLastRow, 4).setValue(borrowerEmail);
  historySheet.getRange(historyLastRow, 5).setValue(borrowerName);
  historySheet.getRange(historyLastRow, 6).setValue(borrowerClass);
  historySheet.getRange(historyLastRow, 7).setValue(action);
  historySheet.getRange(historyLastRow, 8).setValue(dueDate);
  historySheet.getRange(historyLastRow, 9).setValue(userEmail); // Set Processed by in column 9
  historySheet.getRange(historyLastRow, 10).setValue(returnDate); // Set Return Date in column 10
}
