using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToMhtml
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("InputDocument.mht");

        // Find the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply a built‑in table style using its identifier.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally, specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the modified document back to MHTML format.
        doc.Save("OutputDocument.mht");
    }
}
