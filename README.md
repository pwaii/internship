# internship
Open your report in Power BI Desktop.

Go to the Visualizations pane on the right.

Click Get more visuals.
It looks like three dots ... or a small marketplace icon.

Search for Deneb.

Click Add.

Back in the report canvas, click the new Deneb visual icon in the Visualizations pane.

Resize the blank Deneb visual on the page.

In the Data / Fields pane, drag the required fields into the Deneb visual’s field well.
Example for Visual 1, add:
Company
ProductName
IndexName
Currency
CohortNo
ParticipationRate
CapRate
FloorRate
IndexBaseValue
Return From Cohort Base
Participated Return
Credited Return
Risk Score

With the Deneb visual selected, click Create new specification.

Choose Vega-Lite.

Choose Empty specification or Blank.

Open this file:

   [PowerBI_Deneb_Claude_Style_Visuals.txt](C:/Users/legion/Documents/IndexBasedInsurance/PowerBI_Deneb_Claude_Style_Visuals.txt)
Copy one full JSON block.
   For example, copy everything under:
   VISUAL 1 - EXECUTIVE PRODUCT RETURN SCATTER
   starting from:
{
  "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
   through the final closing:
}
Paste that JSON into Deneb’s editor, replacing the existing content.

Click Apply.

The visual should render on the canvas.
