using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load the HTML document that contains a table.
        Document doc = new Document("input.html");

        // Find the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply a built‑in table style by its identifier.
        // For example, use the LightGrid style.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally, specify which parts of the style should be applied.
        // Here we enable first row formatting and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document.
        doc.Save("output.docx");
    }
}
