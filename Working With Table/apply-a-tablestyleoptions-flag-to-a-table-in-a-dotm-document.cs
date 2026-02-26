using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load the existing DOTM template.
        Document doc = new Document("Input.dotm");

        // Find the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply a built‑in table style (optional, can be omitted if the table already has a style).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Set the desired style options using the TableStyleOptions flags.
        // Example: apply formatting to the first row, first column, and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.FirstColumn |
                              TableStyleOptions.RowBands;

        // Save the modified document as a DOTM file.
        doc.Save("Output.dotm");
    }
}
