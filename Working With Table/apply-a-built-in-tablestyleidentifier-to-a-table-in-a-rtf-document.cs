using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToRtf
{
    static void Main()
    {
        // Load the existing RTF document.
        Document doc = new Document("InputDocument.rtf");

        // Assume the document contains at least one table.
        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply a built‑in table style by setting the StyleIdentifier.
        // For example, use the Light Grid style.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally, specify which parts of the style are applied.
        // Here we enable first row, first column and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the modified document back to RTF format.
        doc.Save("OutputDocument.rtf");
    }
}
