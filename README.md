# Shipment Tracking and Comparison System
An Excel Office Script that automates the process of comparing shipment data between two worksheets, identifies missing shipments, and creates organized customer-specific reports.

## Overview
This system helps logistics teams quickly identify which shipments from their main document are missing from their control tracking sheet. It automatically generates individual worksheets for each customer containing their missing shipment details, formatted for easy review and action.

## What This Script Does

### Core Functionality

1. Compares Two Data Sources: Matches shipment IDs between your main shipment document and a control tracking sheet
2. Identifies Missing Shipments: Finds which shipments exist in your main document but are missing from the control sheet
3. Creates Customer-Specific Reports: Generates separate worksheets for each customer containing only their missing shipments
4. Provides Summary Information: Shows you exactly which shipments were found and which were not found

### Automated Processing Steps

1. Cleans and prepares your origin document by removing unnecessary columns
2. Extracts all unique shipment IDs from both worksheets
3. Performs the comparison to identify matches and mismatches
4. Creates a summary section showing the results
5. Generates individual customer worksheets with missing shipment details
6. Applies formatting and formulas for better readability

## Required Worksheet Structure

Your workbook **must contain** these two worksheets:

1. "Origin Document" Worksheet
This contains your main shipment data with these expected columns:
- Shipment ID (required - used for comparison)
- Vessel Name
- Voyage No
- Est Depart Date
- Discharge Location
- Arrival Location
- PO No
- Master BL No
- Container No
- Freight Type
- Date Range

Note: Depending on your business requirements, the script can automatically remove columns like "Ship Mode", "Seal No", "Cartons". To understand how the script works, if you don't know the data, just put a placeholder value.

2. "kontrol" Worksheet

This contains your reference shipment IDs for comparison. The script will:
- Process all cells in this worksheet
- Extract unique shipment IDs
- Use these as the reference list for comparison

# How to Use

## Prerequisites

Excel with Office Scripts enabled (Excel for web, Excel for Windows with Microsoft 365)
Your data organized in the two required worksheets

## Step-by-Step Instructions

1. Prepare Your Data
- Ensure your main shipment data is in a worksheet named "Origin Document"
- Ensure your control/reference data is in a worksheet named "kontrol"
- Make sure shipment IDs are in the first column of your origin document

2. Run the Script
- Open the Office Scripts panel in Excel
- Paste the script code into a new script
- Click "Run" to execute

3. Review the Results
- Check the summary section added to your "Origin Document" worksheet
- Review the individual customer worksheets created for missing shipments

## What Happens When You Run the Script

1. Data Cleaning: The header rows and the columns you specify are removed from your origin document
2. Comparison Process: All shipment IDs are extracted and compared against the shipment IDs in the kontrol sheet.
3. Summary Creation: A table is added to your origin document showing:
- "ShipKeys Found" column: Shipments that exist in both documents
- "ShipKeys NOT Found" column: Shipments missing from the control sheet
4. Customer Worksheets: Individual sheets are created for each customer with missing shipments

## Understanding the Output
### Summary Table
Located at the bottom of your "Origin Document" worksheet:
- Shipment IDs found in the kontrol sheet,
- Shipment IDs not found in the kontrol sheet.

### Customer Worksheets

# Named using the first 3 characters of the shipment ID (assumed to be customer code)
- Formatted with color-coded headers and proper date formatting
  - Green columns: Primary shipment information (Shipment ID, Vessel Name, etc.)
  - Red columns: Secondary reference information (Booking ID, Shipper, etc.)
- Contains only that customer's missing shipment data
- Includes a formula combining Shipment ID and Booking ID

## Customer Code Logic
The script assumes that the first 3 characters of each shipment ID represent the customer code. For example:
- Shipment ID "ABC123456" → Customer "ABC"
- Shipment ID "XYZ789012" → Customer "XYZ"

If your shipment IDs use a different customer identification pattern, you'll need to modify the `getNotFoundCustomers()` function.

## Customization Options

### Modifying Columns to Remove
In the `cleanSheet()` function, you can adjust the columnDeletes array to change which columns are removed:

``` typescript
const columnDeletes = [
  "Ship Mode", 
  "Destination Service Type", 
  "Est Discharge Date", 
  // Add or remove column names as needed
];
```

### Changing Customer Code Logic

Modify this line in `getNotFoundCustomers()` to change how customer codes are extracted:

``` typescript
// Current: uses first 3 characters
notFoundCustomersArr.push(shipKey.slice(0, 3));

// Example: use characters 4-6 instead
notFoundCustomersArr.push(shipKey.slice(3, 6));
```

### Adjusting Output Formatting

In `createNotFoundCustomerSheets()`, you can modify:

- greenBackgoundHeaders: Columns with green background (primary data)
- redBackgroundHeaders: Columns with red background (secondary data)
- tableHeaders: All available column headers

## Troubleshooting

### Common Issues and Solutions

#### Error: "Cannot find worksheet 'Origin Document'"

Ensure your main data worksheet is named exactly "Origin Document"
Check for extra spaces or different capitalization

#### Error: "Cannot find worksheet 'kontrol'"

Ensure your control worksheet is named exactly "kontrol" (lowercase)

#### No customer worksheets created

Verify that some shipments are actually missing from the control sheet
Check that shipment IDs follow the expected format (customer code in first 3 characters)

#### Script runs but no data appears

Ensure your origin document has data starting from the expected row/column positions
Check that shipment IDs are in the first column

#### Performance Considerations

- Large datasets (10,000+ rows) may take several minutes to process
- The script logs progress information to the console for monitoring
- Consider running during off-peak hours for very large datasets

## Technical Details

### Dependencies
- Excel Office Scripts environment
- No external libraries required

### Browser Compatibility
- Works in Excel for web (all modern browsers)
- Compatible with Excel for Windows with Microsoft 365 subscription

### Data Limitations
- Maximum worksheet size limits apply (Excel's standard limits)
- Script memory limitations may affect very large datasets

### Contributing

When contributing to this script:
- Maintain Verbose Comments: All functions should have detailed explanations
- Use Explicit Variable Names: Prefer shipmentKeyArray over arr
- Add Debug Logging: Include console.log statements for troubleshooting
- Test with Sample Data: Verify changes work with realistic datasets

### License

This project is GPL v2.0 open source. Please include attribution when redistributing or modifying.

### Support

For issues or questions:
- Check the troubleshooting section above
- Review the detailed code comments for specific function behavior
- Test with a small sample dataset first
- Ensure your data format matches the expected structure

Last Updated: May 2025
Compatible With: Excel Office Scripts, Microsoft 365
