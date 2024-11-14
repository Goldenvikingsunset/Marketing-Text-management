# PRItemListExt - Page Extension for "Item List"

This page extension adds functionality to the "Item List" page in Business Central, allowing users to import, export, and clear entity text data using Excel files.

## Actions

### Import Entity Text from Excel
- **Caption**: Import Entity Text from Excel
- **Application Area**: All
- **Promoted**: Yes
- **Promoted Category**: Process
- **Promoted Is Big**: Yes
- **Image**: Import

**Trigger OnAction**:
- Reads data from an Excel sheet.
- Imports entity text data into the system.

### Export Entity Text to Excel
- **Caption**: Export Entity Text to Excel
- **Application Area**: All
- **Promoted**: Yes
- **Promoted Category**: Process
- **Promoted Is Big**: Yes
- **Image**: Export

**Trigger OnAction**:
- Exports entity text data from the system to an Excel sheet.

### Clear All Entity Text
- **Caption**: Clear All Entity Text
- **Application Area**: All
- **Promoted**: Yes
- **Promoted Category**: Process
- **Promoted Is Big**: Yes
- **Image**: Delete

**Trigger OnAction**:
- Clears all entity text data from the system.

## Variables
- **TempExcelBuffer**: Temporary record of "Excel Buffer".
- **UploadExcelMsg**: Label 'Please Choose the Excel file.'
- **NoFileFoundMsg**: Label 'No Excel file found!'
- **ExcelImportSuccess**: Label 'Excel is successfully imported.'
- **ExcelExportSuccess**: Label 'Entity Text data has been exported to Excel.'

## Local Procedures

### ImportEntityTextFromExcel
- Imports entity text data from an Excel sheet into the system.

### ExportEntityTextToExcel
- Exports entity text data from the system to an Excel sheet.

### ClearAllEntityText
- Clears all entity text data from the system.

### ReadExcelSheet
- Reads data from an uploaded Excel sheet.

### GetValueAtCell
- Retrieves the value at a specific cell in the Excel sheet.

## Usage

1. **Import Entity Text from Excel**: Click the "Import Entity Text from Excel" action to upload an Excel file and import entity text data.
2. **Export Entity Text to Excel**: Click the "Export Entity Text to Excel" action to export entity text data to an Excel file.
3. **Clear All Entity Text**: Click the "Clear All Entity Text" action to remove all entity text data from the system.

## Notes
- Ensure the Excel file format is correct before importing.
- The system will prompt messages for successful import/export operations or errors if any issues occur.
