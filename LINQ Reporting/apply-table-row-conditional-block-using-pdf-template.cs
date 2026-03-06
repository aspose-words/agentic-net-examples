using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the Word template that contains a table.
        // The template can be any .docx file; Aspose.Words will later save it as PDF.
        Document doc = new Document("Template.docx");

        // Locate the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // ------------------------------------------------------------
        // Create a custom table style that will hold conditional formatting.
        // ------------------------------------------------------------
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Example: Apply a background shading to the last row of the table.
        // This uses the ConditionalStyle collection of the table style.
        customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;

        // Example: Make the first row bold.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Enable the conditional style options on the table.
        // By default only FirstRow, FirstColumn and RowBands are enabled.
        // We need to add LastRow (and any other we used) to the StyleOptions flags.
        table.StyleOptions |= TableStyleOptions.LastRow | TableStyleOptions.FirstRow;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // ------------------------------------------------------------
        // Optionally hide a specific row based on a runtime condition.
        // Here we hide the second row if a condition is met.
        // ------------------------------------------------------------
        bool hideSecondRow = true; // Replace with your actual condition.
        if (hideSecondRow && table.Rows.Count > 1)
        {
            // The Hidden property hides the entire row when the document is rendered.
            table.Rows[1].Hidden = true;
        }

        // ------------------------------------------------------------
        // Save the modified document as a PDF.
        // ------------------------------------------------------------
        doc.Save("Result.pdf", SaveFormat.Pdf);
    }
}
