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
    var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow(), 10).getValues(); // Full Inventory sheet range

    for (var i = 0; i < inventoryData.length; i++) {
      if (inventoryData[i][0] == itemID) { // Match the Item ID in Inventory (Column A)
        
        var currentStatus = inventoryData[i][4]; // Status column in Inventory (Column E)

        // Action: Lend Item
        if (action === 'Lend Item') {
          // Check if the item is on hold
          if (inventoryData[i][9] === 'Yes') { // Check if "On Hold?" column (Column J) is "Yes"
            ui.alert('Item is on Hold. Review "On Hold" and remove the hold before lending.');
            return; // Exit the function to prevent further processing
          }

          // Check if the item is already lent out
          if (currentStatus === 'On Lend') {
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is already on loan.');
            return; // Exit the function to prevent further processing
          }

          // Proceed with lending the item
          inventorySheet.getRange(i + 2, 5).setValue('On Lend'); // Update Status to "On Lend"
          inventorySheet.getRange(i + 2, 7).setValue(borrowerEmail); // Update Current Borrower
          inventorySheet.getRange(i + 2, 8).setValue(new Date()); // Update Loan Date to current date
          
          var dueDate = new Date();
          dueDate.setDate(dueDate.getDate() + 7); // Set Expected Return Date to 1 week from now
          inventorySheet.getRange(i + 2, 9).setValue(dueDate); // Update Expected Return Date

          // Log Lend Item action to History sheet
          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Lend Item', dueDate, userEmail);

        // Action: Return Item
        } else if (action === 'Return Item') {
          // Check if the item is already in stock
          if (currentStatus === 'In Stock') {
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is already in stock.');
            return; // Exit the function to prevent further processing
          }

          // Proceed with returning the item
          inventorySheet.getRange(i + 2, 5).setValue('In Stock'); // Update Status to "In Stock"
          inventorySheet.getRange(i + 2, 7).setValue(''); // Clear Current Borrower
          inventorySheet.getRange(i + 2, 8).setValue(''); // Clear Loan Date
          inventorySheet.getRange(i + 2, 9).setValue(''); // Clear Expected Return Date
          
          // Log Return Item action to History sheet
          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Return Item', '', userEmail);

        // Action: Hold Item
        } else if (action === 'Hold Item') {
          // Check if the item is already on hold
          if (inventoryData[i][9] === 'Yes') { // Check if "On Hold?" column (Column J) is "Yes"
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is already on hold.');
            return; // Exit the function to prevent further processing
          }

          // Proceed with placing the item on hold
          inventorySheet.getRange(i + 2, 10).setValue('Yes'); // Set "On Hold?" column to "Yes"

          // Insert record into On Hold sheet
          var onHoldLastRow = onHoldSheet.getLastRow() + 1;
          onHoldSheet.getRange(onHoldLastRow, 1).setValue(itemID); // Set Item ID
          onHoldSheet.getRange(onHoldLastRow, 2).setValue(itemName); // Set Item Name
          onHoldSheet.getRange(onHoldLastRow, 3).setValue(borrowerEmail); // Set Email
          onHoldSheet.getRange(onHoldLastRow, 4).setValue(borrowerName); // Set Name
          onHoldSheet.getRange(onHoldLastRow, 5).setValue(borrowerClass); // Set Class
          onHoldSheet.getRange(onHoldLastRow, 6).setValue(new Date()); // Set Date Requested to current date

          // Log Hold Item action to History sheet
          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Hold Item', '', userEmail);

        // Action: Remove Hold
        } else if (action === 'Remove Hold') {
          // Check if the item is on hold
          if (inventoryData[i][9] !== 'Yes') { // Check if "On Hold?" column (Column J) is not "Yes"
            ui.alert('Error: The item "' + inventoryData[i][1] + '" is not currently on hold.');
            return; // Exit the function to prevent further processing
          }

          // Proceed with removing the hold
          inventorySheet.getRange(i + 2, 10).setValue('No'); // Set "On Hold?" column to "No"

          // Find and remove the record from the On Hold sheet
          var onHoldData = onHoldSheet.getRange(2, 1, onHoldSheet.getLastRow(), 6).getValues(); // Full On Hold sheet range
          for (var j = 0; j < onHoldData.length; j++) {
            if (onHoldData[j][0].toString().trim() === itemID) { // Match the Item ID in On Hold sheet
              onHoldSheet.deleteRow(j + 2); // Delete the row in the On Hold sheet
              break; // Exit the loop once the record is found and deleted
            }
          }

          // Log Remove Hold action to History sheet
          logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, 'Remove Hold', '', userEmail);
        }

        // After updating, clear the Action and Borrower fields in Lending sheet
        lendingSheet.getRange(row, 6).setValue(''); // Clear Action dropdown
        lendingSheet.getRange(row, 3).setValue(''); // Clear Borrower Email in Lending
        lendingSheet.getRange(row, 4).setValue(''); // Clear Borrower Name in Lending
        lendingSheet.getRange(row, 5).setValue(''); // Clear Borrower Class in Lending

        break; // Exit the loop after the match is found
      }
    }
  }
}

// Function to log actions to the History sheet
function logHistoryAction(itemID, itemName, borrowerEmail, borrowerName, borrowerClass, action, dueDate, userEmail) {
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  var historyLastRow = historySheet.getLastRow() + 1;
  
  historySheet.getRange(historyLastRow, 1).setValue(new Date()); // Set Date of the entry
  historySheet.getRange(historyLastRow, 2).setValue(itemID); // Set Item ID
  historySheet.getRange(historyLastRow, 3).setValue(itemName); // Set Item Name
  historySheet.getRange(historyLastRow, 4).setValue(borrowerEmail); // Set Email
  historySheet.getRange(historyLastRow, 5).setValue(borrowerName); // Set Name
  historySheet.getRange(historyLastRow, 6).setValue(borrowerClass); // Set Class
  historySheet.getRange(historyLastRow, 7).setValue(action); // Set Action (Lend Item, Return Item, Hold Item, Remove Hold)
  historySheet.getRange(historyLastRow, 8).setValue(dueDate); // Set Due Date (if applicable)
  historySheet.getRange(historyLastRow, 9).setValue(userEmail); // Set Processed by (User who clicked Finalize)
}
