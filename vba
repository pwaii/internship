dd these actions:
Excel Online (Business) – Run script
Location: your SharePoint site
Library: your document library
File: Identifier from the trigger
Script: your Pivot Builder script

Parse JSON
Content: result from Run script
Schema:

{
  "type": "object",
  "properties": {
    "action": { "type": "string" },
    "outputFileName": { "type": "string" },
    "originalFileName": { "type": "string" },
    "dataSheet": { "type": "string" },
    "selectedTemplate": { "type": "string" }
  },
  "required": [
    "action",
    "outputFileName",
    "originalFileName",
    "dataSheet",
    "selectedTemplate"
  ]
}
SharePoint – Get file content
File Identifier: Identifier from the trigger

SharePoint – Create file
Folder Path: your output folder
File Name: outputFileName from Parse JSON
File Content: output from Get file content

Delay
Count: 5
Unit: Seconds

Excel Online (Business) – Run script
File: Identifier from Create file
Script: the same Pivot Builder script

If the second Run script rejects that identifier, insert Get file metadata using path after Create file and use its Identifier.
Run It
In PivotBuilder_Setup, enter the same filename in Output File Name for all selected rows.
Select Create output copy in Action.
Save the workbook.
In SharePoint, select the raw workbook.
Open Automate/Power Automate > Create Pivot Output.
Run the flow.
Use a new filename each time; the flow should not overwrite an existing report.
