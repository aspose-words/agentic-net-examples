using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load an existing WORDML document.
        Document doc = new Document("Input.docx"); // Replace with the actual WORDML file path.

        // Find the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply desired style options to the table.
        // Example: apply first row formatting, row banding, and first column formatting.
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.RowBands |
                              TableStyleOptions.FirstColumn;

        // Optionally, set a built‑in style identifier for the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Save the modified document.
        doc.Save("Output.docx"); // Replace with the desired output path.
    }
}
