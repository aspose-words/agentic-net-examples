using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToMhtml
{
    static void Main()
    {
        // Load the MHTML document.
        string inputPath = "input.mhtml";
        Document doc = new Document(inputPath);

        // Find the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table != null)
        {
            // Apply a built‑in table style by its identifier.
            table.StyleIdentifier = StyleIdentifier.LightGrid;

            // Optionally specify which parts of the style are applied.
            table.StyleOptions = TableStyleOptions.FirstRow |
                                 TableStyleOptions.RowBands |
                                 TableStyleOptions.FirstColumn;
        }

        // Save the modified document back to MHTML.
        string outputPath = "output.mhtml";
        doc.Save(outputPath);
    }
}
