function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Check if we're on the Lending sheet
  if (sheet.getName() === 'Lending') {
    
    var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory'); // Reference Inventory sheet
    var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow(), 2).getValues(); // Get Item ID and Item Name from Inventory
    
    // Column A is Item ID, Column B is Item Name in Lending sheet
    var itemIDCell = sheet.getRange(range.getRow(), 1); // Item ID
    var itemNameCell = sheet.getRange(range.getRow(), 2); // Item Name
    
    // If Item ID is edited (Column A)
    if (range.getColumn() == 1 && itemIDCell.getValue() !== "") {
      var itemID = itemIDCell.getValue();
      
      // Find the corresponding Item Name from Inventory
      for (var i = 0; i < inventoryData.length; i++) {
        if (inventoryData[i][0] == itemID) {
          itemNameCell.setValue(inventoryData[i][1]); // Set the Item Name in Column B
          break;
        }
      }
    }
    
    // If Item Name is edited (Column B)
    if (range.getColumn() == 2 && itemNameCell.getValue() !== "") {
      var itemName = itemNameCell.getValue();
      
      // Find the corresponding Item ID from Inventory
      for (var i = 0; i < inventoryData.length; i++) {
        if (inventoryData[i][1] == itemName) {
          itemIDCell.setValue(inventoryData[i][0]); // Set the Item ID in Column A
          break;
        }
      }
    }
  }
}