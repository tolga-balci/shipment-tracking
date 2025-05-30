/**
 * SHIPMENT TRACKING AND COMPARISON SYSTEM
 * 
 * This Excel Office Script compares shipment data between two worksheets:
 * - "Origin Document": Contains the main shipment data to be processed
 * - "kontrol": Contains reference shipment IDs for comparison
 * 
 * The script identifies which shipments are found/not found in the control sheet,
 * then creates separate worksheets for each customer with missing shipments.
 * 
 * Main workflow:
 * 1. Clean and prepare the origin document data
 * 2. Extract shipment keys from both sheets
 * 3. Compare the keys to find matches and mismatches
 * 4. Create summary information for the user
 * 5. Generate individual customer worksheets for missing shipments
 */

/**
 * MAIN FUNCTION - Entry point of the script
 * 
 * This function orchestrates the entire shipment comparison and reporting process.
 * It coordinates data extraction, comparison, and worksheet generation.
 * 
 * @param workbook - The Excel workbook containing all worksheets to process
 */
function main(workbook: ExcelScript.Workbook) {
  // Step 1: Get references to the two main worksheets we'll be working with
  // "Origin Document" contains the shipment data we want to analyze
  let Sheet = workbook.getWorksheet("Origin Document");
  
  // "kontrol" contains the reference shipment IDs we'll compare against
  let kontrolSheet = workbook.getWorksheet("kontrol");
  
  // Get the range of cells that actually contain data (not empty cells)
  const usedRange = sheet.getUsedRange();
  
  // Declare variable to hold the main data table
  let originTable: ExcelScript.Table;

  // Step 2: Create or get the main data table
  // Check if a table named 'originTable' already exists
  if (sheet.getTable('originTable')) {
    // If table exists, use the existing one
    originTable = sheet.getTable('originTable');
  } else {
    // If no table exists, we need to create one
    // First clean the sheet by removing unwanted rows and columns
    cleanSheet(sheet);
    
    // Create a new table from the cleaned data range
    // The 'true' parameter indicates the first row contains headers
    originTable = sheet.addTable(usedRange, true);
    
    // Give the table a specific name so we can reference it later
    originTable.setName("originTable");
  }

  // Step 3: Initialize arrays to store shipment keys (IDs)
  // These will hold the shipment IDs from both worksheets
  let shipKeys: string[] = [];           // Shipment IDs from origin document
  let kontrolShipKeys: string[] = [];    // Shipment IDs from control sheet

  // Remove any existing auto-filter to ensure clean data processing
  originTable.getAutoFilter().remove();
  
  // Step 4: Extract shipment keys from both worksheets
  // Get all unique shipment IDs from the main document
  getShipKeys(usedRange, shipKeys);

  // Get all unique shipment IDs from the control sheet
  let kontrolSheetDistinctShipKeys = getKontrolDistinctShipKeys(usedRange, kontrolSheet);
  
  // Log the counts for debugging and monitoring
  console.log(`fn main Shipkeys: ${shipKeys.length}`);
  console.log(`fn main kontrolDistinctShipKeys: ${kontrolSheetDistinctShipKeys.length}`);

  // Step 5: Compare the shipment keys to find matches and mismatches
  // This returns two arrays: [found keys, not found keys]
  const [foundShipKeys, notFoundShipKeys] = compareShipKeys(shipKeys, kontrolSheetDistinctShipKeys);

  // Log the results of the comparison
  console.log(`fn main notFoundShipKeys: ${notFoundShipKeys}`);

  // Step 6: Create a summary section on the original sheet
  // This shows the user which shipments were found/not found
  updateUserWithShipKeys(sheet, usedRange, foundShipKeys, notFoundShipKeys);

  // Step 7: Process the shipments that weren't found
  // Extract unique customer IDs from the missing shipment keys
  // (assumes first 3 characters of shipment ID represent customer)
  let notFoundCustomersList: string[] = getNotFoundCustomers(notFoundShipKeys);
  
  // Get detailed shipment information for all missing shipments
  let notFoundCustomersShipmentDetails: string[][] = getNotFoundCustomersShipmentDetails(originTable, notFoundShipKeys);
  console.log(`fn main notFoundCustomersShipmentDetails: ${notFoundCustomersShipmentDetails}`)

  // Step 8: Create individual worksheets for each customer with missing shipments
  // Each worksheet will contain only that customer's missing shipment data
  createNotFoundCustomerSheets(workbook, notFoundCustomersList, notFoundCustomersShipmentDetails);
}

/**
 * SHEET CLEANING FUNCTION
 * 
 * This function prepares the origin document by removing unnecessary data
 * and columns that aren't needed for the comparison process.
 * 
 * @param sheet - The worksheet to clean up
 */
function cleanSheet(sheet: ExcelScript.Worksheet) {
  // Step 1: Remove the first 4 rows (likely headers or metadata we don't need)
  // DeleteShiftDirection.up means remaining rows move up to fill the gap
  sheet.getRange("1:4").delete(ExcelScript.DeleteShiftDirection.up);
  
  // Step 2: Define columns that should be removed from the data
  // These are columns that aren't relevant for your shipment comparison
  const columnDeletes = [
    "Ship Mode", 
    "Destination Service Type", 
    "Est Discharge Date", 
    "Item No", 
    "Seal No", 
    "Cartons", 
    "Units", 
    "Volume", 
    "Weight", 
    "Ci Last Modified", 
    "Pl Last Modified", 
    "Certificate Required", 
    "Commercial Invoice No", 
    "FCR No", 
    "F&W"
  ];

  // Step 3: Create a temporary table to work with the data
  // This makes it easier to manipulate columns by name
  let tempUsedRange = sheet.getUsedRange();
  let tempUsedRangeTable = sheet.addTable(tempUsedRange, true);

  // Step 4: Remove each unwanted column
  columnDeletes.forEach(columnName => {
    // Try to find the column by name
    const column = tempUsedRangeTable.getColumnByName(columnName);
    
    // If the column exists, delete it
    if (column) {
      column.delete();
    }
  });

  // Step 5: Format the remaining data for better readability
  // Auto-fit column widths to content
  tempUsedRangeTable.getRange().getFormat().autofitColumns();
  
  // Auto-fit row heights to content
  tempUsedRangeTable.getRange().getFormat().autofitRows();
  
  // Convert back to a regular range (remove table formatting)
  tempUsedRangeTable.convertToRange();
}

/**
 * SHIPMENT KEY EXTRACTION FUNCTION
 * 
 * This function extracts all unique shipment IDs from the origin document.
 * It assumes the shipment ID is in the first column of the data.
 * 
 * @param usedRange - The range of cells containing data
 * @param shipKeys - Array to populate with unique shipment keys (passed by reference)
 * @returns Array of unique shipment keys
 */
function getShipKeys(usedRange: ExcelScript.Range, shipKeys: string[]): string[] {
  // Get the total number of rows in the data range
  const rowCount = usedRange.getRowCount();
  
  // Start from row 1 (row 0 contains headers)
  const firstDataRow = 1;
  
  // Shipment ID is assumed to be in the first column (index 0)
  const shipKeyColumnIndex = 0;

  // Temporary array to collect all shipment keys (including duplicates)
  let shipKeysArr: string[] = [];

  // Step 1: Loop through each data row to extract shipment IDs
  for (let i = firstDataRow; i < rowCount; i++) {
    // Get the value from the shipment ID column
    const shipKey = usedRange.getCell(i, shipKeyColumnIndex).getValue()?.toString();
    
    // Only add non-empty values
    if (shipKey) {
      shipKeysArr.push(shipKey);
    }
  }
  
  // Step 2: Remove duplicates using a Set
  // Set automatically eliminates duplicate values
  const shipKeysSet = new Set(shipKeysArr);
  
  // Log counts for debugging
  console.log(`fn getShipKeys shipKeysArr Length: ${shipKeysArr.length}`)
  
  // Step 3: Clear the original array and populate with unique values
  shipKeys.length = 0; // Clear the original array completely
  shipKeys.push(...Array.from(shipKeysSet)); // Add all unique values
  
  // Log final counts for verification
  console.log(`fn getShipKeys shipKeysSet Length: ${shipKeysSet.size}`)
  console.log(`fn getShipKeys shipKeys Length: ${shipKeys.length}`)
  
  return shipKeys;
}

/**
 * CONTROL SHEET KEY EXTRACTION FUNCTION
 * 
 * This function extracts all unique shipment IDs from the control worksheet.
 * It processes all cells in the sheet, flattens the data, and removes duplicates.
 * 
 * @param usedRange - The range from origin sheet (not used in current implementation)
 * @param kontrolSheet - The control worksheet containing reference shipment IDs
 * @returns Array of unique shipment keys from the control sheet
 */
function getKontrolDistinctShipKeys(usedRange: ExcelScript.Range, kontrolSheet: ExcelScript.Worksheet): string[] {
 
  // Step 1: Get the data range (skip header row by using offset)
  // getOffsetRange(1, 0) means start 1 row down from the used range
  const kontrolSheetDataRange = kontrolSheet.getUsedRange().getOffsetRange(1, 0);

  // Step 2: Get all cell values - this returns a 2D array
  // Each row is an array, so we have an array of arrays
  // Cell values can be string, number, or boolean
  const rawCellValues: (string | number | boolean)[][] = kontrolSheetDataRange.getValues();

  // Step 3: Flatten the 2D array into a 1D array
  // We use reduce with concat to combine all rows into one array
  // This takes [[a,b],[c,d]] and makes it [a,b,c,d]
  const flattenedValues: (string | number | boolean)[] = rawCellValues.reduce(
    (accumulator, currentRow) => accumulator.concat(currentRow),
    []
  );

  // Step 4: Convert each value to string and filter out empty/null values
  const stringValues: string[] = flattenedValues
    .map(cellValue => {
      // Convert each cell value to string explicitly
      if (cellValue === null || cellValue === undefined) {
        return ""; // Convert null/undefined to empty string
      }
      return String(cellValue); // Convert number/boolean/string to string
    })
    .filter(stringValue => stringValue.trim() !== ""); // Remove empty strings after trimming whitespace

  // Step 5: Remove duplicates using Set
  // Set automatically eliminates duplicate values
  const uniqueStringValues: Set<string> = new Set(stringValues);

  // Step 6: Convert Set back to array for easier handling
  const kontrolShipKeys: string[] = Array.from(uniqueStringValues);

  // Step 7: Log for debugging and monitoring
  console.log(`fn getKontrolDistinctShipKeys: Found ${kontrolShipKeys.length} unique Shipment IDs`);

  return kontrolShipKeys;
}

/**
 * SHIPMENT KEY COMPARISON FUNCTION
 * 
 * This function compares shipment keys from the origin document against
 * the control sheet to determine which ones exist and which are missing.
 * 
 * @param shipKeys - Array of shipment IDs from the origin document
 * @param kontrolSheetDistinctShipKeys - Array of shipment IDs from control sheet
 * @returns Array containing two arrays: [foundShipKeys, notFoundShipKeys]
 */
function compareShipKeys(shipKeys: string[], kontrolSheetDistinctShipKeys: string[]): string[][] {
  
  // Step 1: Create a Set from control sheet keys for efficient lookup
  // Set.has() is much faster than Array.includes() for large datasets
  const kontrolDistinctShipKeysSet = new Set(Array.from(kontrolSheetDistinctShipKeys))
  
  // Convert back to array (this step seems redundant in current code)
  const kontrolDistinctShipKeysArr = Array.from(kontrolDistinctShipKeysSet)

  /* Alternative approach using traditional loop (commented out):
  const foundShipKeys: string[] = [];
  const notFoundShipKeys: string[] = [];
  
  for (const shipKey of shipKeys) {
    if (kontrolDistinctShipKeysSet.has(shipKey)) {
      foundShipKeys.push(shipKey);
    } else {
      notFoundShipKeys.push(shipKey);
    }
  }
  */
  
  // Debug logging to understand the data structure
  console.log(`fn compareShipKeys kontrolSheetDistinctShipKeys: ${kontrolSheetDistinctShipKeys}`)
  console.log(`fn compareShipKeys kontrolDistinctShipKeysArr: ${typeof kontrolDistinctShipKeysArr}`)
  
  // Step 2: Use array filter methods to separate found and not found keys
  // Filter creates new arrays containing only elements that match the condition
  
  // Found keys: shipment IDs that exist in the control sheet
  const foundShipKeys = shipKeys.filter(shipKey => kontrolDistinctShipKeysArr.includes(shipKey));
  
  // Not found keys: shipment IDs that do NOT exist in the control sheet
  const notFoundShipKeys = shipKeys.filter(shipKey => !kontrolDistinctShipKeysArr.includes(shipKey));

  // Debug logging to verify the results
  console.log(`fn compareShipKeys: ${kontrolDistinctShipKeysSet}`)
  console.log(`fn compareShipKeys: ${foundShipKeys}`)
  console.log(`fn compareShipKeys: ${notFoundShipKeys}`)

  // Step 3: Return both arrays as a single array
  // This allows the caller to destructure: [found, notFound] = compareShipKeys(...)
  return [foundShipKeys, notFoundShipKeys];
}

/**
 * USER SUMMARY UPDATE FUNCTION
 * 
 * This function creates a summary section on the origin worksheet showing
 * the user which shipment keys were found and which were not found.
 * 
 * @param sheet - The worksheet to update with summary information
 * @param usedRange - The range of existing data
 * @param foundShipKeys - Array of shipment keys that were found in control sheet
 * @param notFoundShipKeys - Array of shipment keys that were NOT found in control sheet
 */
function updateUserWithShipKeys(sheet: ExcelScript.Worksheet, usedRange: ExcelScript.Range, foundShipKeys: string[], notFoundShipKeys: string[]) {
  // Log the arrays for debugging
  console.log(`fn updateUserWithShipKeys foundShipKeys: ${foundShipKeys}`);
  console.log(`fn updateUserWithShipKeys notFoundShipKeys: ${notFoundShipKeys}`);

  // Step 1: Calculate where to place the summary section
  // Find the last row with data
  const tempUsedCellLastRow = usedRange.getLastCell().getRowIndex();
  
  // Get the counts of found and not found keys
  const foundShipKeysLength = foundShipKeys.length;
  const notFoundShipKeysLength = notFoundShipKeys.length;

  // Calculate how many rows we need for the summary table
  // Use the larger of the two arrays to determine table height
  const infoTableRowCount: number = Math.max(foundShipKeysLength, notFoundShipKeysLength);

  console.log(`fn updateUserWithShipKeys infoTableRowCount: ${infoTableRowCount}`);

  // Step 2: Define the position for the summary table
  // Place it 3 rows below the existing data
  const userInfoTableStartRow = tempUsedCellLastRow + 3;
  const userInfoTableStartDataRow = userInfoTableStartRow + 1;  // First data row (after headers)
  const userInfoTableStartColumn = 1;  // Start in column B (index 1)

  // Step 3: Create the column headers for the summary table
  sheet.getCell(userInfoTableStartRow, userInfoTableStartColumn).setValue("ShipKeys Found");
  sheet.getCell(userInfoTableStartRow, userInfoTableStartColumn + 1).setValue("ShipKeys NOT Found");

  // Step 4: Populate the "Found" column if there are found keys
  if (foundShipKeys.length > 0) {
    // Convert array of strings to 2D array format required by Excel
    // Each string becomes a single-element array: ["key1"] -> [["key1"], ["key2"], ...]
    const foundShipKeyValues: string[][] = foundShipKeys.map(key => [key]);
    
    // Define the range where found keys will be written
    const foundShipKeyRange = sheet.getRangeByIndexes(
      userInfoTableStartDataRow,           // Start row
      userInfoTableStartColumn,            // Start column  
      foundShipKeys.length,                // Number of rows
      1                                    // Number of columns
    );
    
    // Write all the found keys at once
    foundShipKeyRange.setValues(foundShipKeyValues);
  }

  // Step 5: Populate the "Not Found" column if there are missing keys
  if (notFoundShipKeys.length > 0) {
    // Convert array of strings to 2D array format
    const notFoundShipKeyValues: string[][] = notFoundShipKeys.map(key => [key]);
    
    // Define the range where not found keys will be written (next column over)
    const notFoundShipKeyRange = sheet.getRangeByIndexes(
      userInfoTableStartRow + 1,           // Start row (after header)
      userInfoTableStartColumn + 1,        // Start column (second column)
      notFoundShipKeys.length,             // Number of rows
      1                                    // Number of columns
    );
    
    // Write all the not found keys at once
    notFoundShipKeyRange.setValues(notFoundShipKeyValues);
  }
}

/**
 * CUSTOMER EXTRACTION FUNCTION
 * 
 * This function extracts unique customer IDs from shipment keys.
 * It assumes the first 3 characters of each shipment ID represent the customer code.
 * 
 * @param notFoundShipKeys - Array of shipment keys that weren't found in control sheet
 * @returns Array of unique customer IDs
 */
function getNotFoundCustomers(notFoundShipKeys: string[]): string[] {
  // Array to collect customer IDs (may include duplicates initially)
  const notFoundCustomersArr: string[] = [];

  // Step 1: Extract customer ID from each shipment key
  notFoundShipKeys.forEach(shipKey => {
    // Take the first 3 characters as the customer identifier
    // Example: "ABC123456" -> "ABC"
    notFoundCustomersArr.push(shipKey.slice(0, 3));
  });

  // Step 2: Remove duplicates to get unique customer list
  // This is important because we'll create one worksheet per customer
  const notFoundCustomersSet = new Set(notFoundCustomersArr);
  const notFoundCustomersList = Array.from(notFoundCustomersSet)

  // Log the unique customer list for debugging
  console.log(`fn createCustomerWorksheets  notFoundCustomersList: ${notFoundCustomersList}\n`);

  return notFoundCustomersList;
}

/**
 * SHIPMENT DETAILS EXTRACTION FUNCTION
 * 
 * This function retrieves the complete shipment details for all shipments
 * that were not found in the control sheet. It returns full row data for each.
 * 
 * @param originTable - The main data table containing all shipment information
 * @param notFoundShipKeys - Array of shipment keys that weren't found
 * @returns 2D array where each sub-array represents a complete shipment record
 */
function getNotFoundCustomersShipmentDetails(originTable: ExcelScript.Table, notFoundShipKeys: string[]): string[][] {
  // Array to store complete shipment details
  const notFoundCustomersShipmentDetails: string[][] = [];
  
  // Step 1: Get the shipment ID column and its values for efficient lookup
  const shipKeyColumn = originTable.getColumn("Shipment ID");
  const shipKeyValues: string[][] = shipKeyColumn.getRangeBetweenHeaderAndTotal().getValues() as string[][];
  
  // Get the full table range and column count for data extraction
  const tableRange = originTable.getRangeBetweenHeaderAndTotal();
  const columnCount = tableRange.getColumnCount();

  // Step 2: Remove duplicates from notFoundShipKeys to avoid processing same shipment multiple times
  const uniqueNotFoundShipKeys = [...new Set(notFoundShipKeys)];

  // Step 3: Create a lookup map for efficient shipment ID to row index mapping
  // This avoids nested loops and improves performance for large datasets
  const shipKeyToRowIndex = new Map<string, number>();
  shipKeyValues.forEach((row, index) => {
    // Map each shipment ID to its row index in the table
    shipKeyToRowIndex.set(row[0] as string, index);
  });

  // Step 4: Process each unique shipment key that wasn't found
  uniqueNotFoundShipKeys.forEach(notFoundShipKey => {
    // Look up the row index for this shipment key
    const rowIndex = shipKeyToRowIndex.get(notFoundShipKey);
    
    if (rowIndex !== undefined) {
      // Row was found, so extract all column data for this shipment
      const rowData: string[] = [];
      
      // Step 5: Extract all cell values from this row
      for (let j = 0; j < columnCount; j++) {
        const cellValue = tableRange.getCell(rowIndex, j).getValue();
        // Convert null values to empty strings and everything else to strings
        rowData.push(cellValue === null ? "" : String(cellValue));
      }
      
      // Add this complete shipment record to our results
      notFoundCustomersShipmentDetails.push(rowData);
    }
  });

  return notFoundCustomersShipmentDetails;
}

/**
 * CUSTOMER WORKSHEET CREATION FUNCTION
 * 
 * This function creates individual worksheets for each customer who has
 * shipments that weren't found in the control sheet. Each worksheet contains
 * only that customer's missing shipment data with proper formatting.
 * 
 * Business requirement is to identify columns in red and green to indicate
 * which column will be filled by whom.
 *
 * @param workbook - The Excel workbook to add worksheets to
 * @param notFoundCustomersList - Array of unique customer IDs
 * @param notFoundCustomersShipmentDetails - 2D array of complete shipment records
 */
function createNotFoundCustomerSheets(workbook: ExcelScript.Workbook, notFoundCustomersList: string[], notFoundCustomersShipmentDetails: string[][]): void {

  console.log(`fn createNotFoundCustomerSheets notFoundCustomersShipmentDetails: ${notFoundCustomersShipmentDetails}`);

  // Step 1: Define the column structure for customer worksheets
  // All possible headers that might be needed
  let tableHeaders: string[] = [
    "Shipment ID", "Booking ID", "Shipment ID - Booking ID", "Shipper", 
    "Vessel Name", "Voyage No", "Est Depart Date", "Place of Origin", 
    "Discharge Location", "House BL No", "Master BL No", "Container No", 
    "Container Size and Type", "Freight Type", "Agent", "Export or Import", 
    "Date Range", "Purchase Order ID", "Sea or Air", "Consignee", 
    "Port of Loading", "Port of Discharge", "Customer"
  ];
  
  // Headers that should have green background (important/primary data)
  let greenBackgoundHeaders: string[] = [
    "Shipment ID", "Vessel Name", "Voyage No", "Est Depart Date", 
    "Place of Origin", "Discharge Location", "Master BL No", "Container No", 
    "Freight Type", "Date Range"
  ];
  
  // Headers that should have red background (secondary/reference data)
  let redBackgroundHeaders: string[] = [
    "Booking ID", "Shipment ID - Booking ID", "Shipper", "Place of Origin", 
    "House BL No", "Container Size and Type", "Export or Import", 
    "Purchase Order ID", "Sea or Air", "Consignee", 
    "Port of Loading", "Port of Discharge", "Customer"
  ];

  // Step 2: Define which columns from the source data we want to display
  // This maps to the actual data structure from the origin table
  const shipmentDataColumns = [
    "Shipment ID", "Vessel Name", "Voyage No", "Est Depart Date",
    "Discharge Location", "Arrival Location", "PO No", "Master BL No",
    "Container No", "Freight Type", "Date Range"
  ];

  // Step 3: Create a worksheet for each customer
  notFoundCustomersList.forEach(customerID => {
    
    // Step 3a: Try to create the worksheet
    try {
      let worksheet = workbook.addWorksheet(String(customerID));
    }
    catch (error) {
      // If worksheet creation fails (e.g., name already exists), skip this customer
      console.log(`Error creating sheet ${customerID}: ${error}`);
      return; // Skip the rest of the processing for this customer.
    }

    // Step 3b: Get reference to the newly created worksheet
    const worksheet = workbook.getWorksheet(customerID);
    if (!worksheet) return; // Safety check: handle case where sheet wasn't created

    // Step 4: Create the header row with proper formatting
    let headerColumnMap = new Map<string, number>(); // Map to store header name -> column index
    
    tableHeaders.forEach((header, index) => {
      const cell = worksheet.getCell(0, index);
      cell.setValue(header);

      // Apply color coding based on header importance
      if (greenBackgoundHeaders.includes(header)) {
        cell.getFormat().getFill().setColor("#00FF00"); // Green background
        headerColumnMap.set(header, index); // Store column index for later data population
      } else if (redBackgroundHeaders.includes(header)) {
        cell.getFormat().getFill().setColor("#FF0000"); // Red background
      }
    });
    
    // Step 5: Populate data rows for this customer
    let rowCounter = 1; // Start populating data from row 1 (row 0 has headers)
    
    notFoundCustomersShipmentDetails.forEach(shipment => {
      // Check if this shipment belongs to the current customer
      // Compare first 3 characters of shipment ID with customer ID
      const customerPrefix = shipment[0].slice(0, 3); // Assuming Shipment ID is the first element
      
      if (customerPrefix === customerID) {
        let dataWritten = false;
        
        // Create an object to map column names to shipment data
        const shipmentData: { [key: string]: string } = {};

        // Debug logging for data mapping verification
        console.log("All shipment data:");
        for (const [key, value] of Object.entries(shipmentData)) {
          console.log(`${key}: ${value}`);
        }

        // Step 5a: Map the shipment data to column names
        shipmentDataColumns.forEach((columnName, index) => {
          shipmentData[columnName] = shipment[index];
          console.log(`fn createNotFoundCustomerSheets: shipmentData[columnName] ${shipmentData[columnName]}, shipment[index]: ${shipment[index]}`);
        });

        // Step 5b: Write data only to green background columns (primary data)
        greenBackgoundHeaders.forEach(header => {
          const columnIndex = headerColumnMap.get(header);
          if (columnIndex !== undefined) {
            worksheet.getCell(rowCounter, columnIndex).setValue(shipmentData[header]);
            console.log(`fn createNotFoundCustomerSheets: Writing ${shipmentData[header]}`)
            dataWritten = true;
          }
        });
        
        // Only increment row counter if we actually wrote data
        if (dataWritten) {
          rowCounter++;
        }
      }
    });
    
    // Step 6: Apply formulas and formatting to the completed worksheet
    insertFormulasToWorksheet(worksheet);
    formatWorksheet(worksheet);
  });
}

/**
 * FORMULA INSERTION FUNCTION
 * 
 * This function adds Excel formulas to customer worksheets.
 * Currently adds a concatenation formula to combine Shipment ID and Booking ID.
 * 
 * @param worksheet - The worksheet to add formulas to
 */
function insertFormulasToWorksheet(worksheet: ExcelScript.Worksheet): void {
  
  // Step 1: Convert the worksheet data to a table for easier column manipulation
  let sheetUsedRange = worksheet.getUsedRange();
  let sheetUsedRangeTable = worksheet.addTable(sheetUsedRange, true);

  // Step 2: Get references to the columns we need for the formula
  const shipKeyColumn = sheetUsedRangeTable.getColumnByName("Shipment ID");
  const bookingKeyColumn = sheetUsedRangeTable.getColumnByName("Booking ID");
  const shipKeyBookingKeyColumn = sheetUsedRangeTable.getColumnByName("Shipment ID - Booking ID");
  
  // Step 3: Define the concatenation formula
  // This formula combines the values from columns A and B with " - " separator
  // Example: If A2="SHIP123" and B2="BOOK456", result will be "SHIP123 - BOOK456"
  const concatenateFormula = "=CONCATENATE(A2,\" - \",B2)";

  // Step 4: Apply the formula to all data rows in the combined column
  shipKeyBookingKeyColumn.getRangeBetweenHeaderAndTotal().setFormula(concatenateFormula)

  // Step 5: Convert back to regular range (remove table structure)
  sheetUsedRangeTable.convertToRange();
}

/**
 * WORKSHEET FORMATTING FUNCTION
 * 
 * This function applies consistent formatting to customer worksheets including
 * date formatting, borders, text wrapping, and column sizing.
 * 
 * @param worksheet - The worksheet to format
 */
function formatWorksheet(worksheet: ExcelScript.Worksheet): void {

  // Step 1: Define which columns contain dates that need special formatting
  const dateColumns = ["Est Depart Date", "Date Range"];

  // Step 2: Create a temporary table to work with column formatting
  let sheetUsedRange = worksheet.getUsedRange();
  let sheetUsedRangeTable = worksheet.addTable(sheetUsedRange, true);

  // Step 3: Apply date formatting to date columns
  dateColumns.forEach(columnName => {
    const column = sheetUsedRangeTable.getColumnByName(columnName)
    if (column) {
      // Set number format to display dates as day.month.year (e.g., 15.03.2024)
      column.getRangeBetweenHeaderAndTotal().setNumberFormat("d.m.yyyy")
    }
  });

  // Step 4: Convert table back to regular range for border formatting
  sheetUsedRangeTable.convertToRange();

  // Step 5: Apply comprehensive border formatting to make the data easier to read
  // All of these border settings create a complete grid around and within the data
  
  // Left edge border - creates border on the left side of the entire data range
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
  
  // Top edge border - creates border on the top of the entire data range
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
  
  // Bottom edge border - creates border on the bottom of the entire data range
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
  
  // Right edge border - creates border on the right side of the entire data range
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
  
  // Inside vertical borders - creates vertical lines between all columns
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
  
  // Inside horizontal borders - creates horizontal lines between all rows
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
  worksheet.getUsedRange().getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

  // Step 6: Apply text formatting for better readability
  
  // Enable text wrapping - allows long text to wrap within cells instead of overflowing
  worksheet.getUsedRange().getFormat().setWrapText(true);
  
  // Set vertical alignment to top - ensures text starts at the top of each cell
  // This is especially useful when cells have different heights due to text wrapping
  worksheet.getUsedRange().getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  
  // Step 7: Auto-fit columns to content
  // This automatically adjusts column widths to fit the content properly
  // Ensures all data is visible without manual column resizing
  worksheet.getUsedRange().getFormat().autofitColumns();
}