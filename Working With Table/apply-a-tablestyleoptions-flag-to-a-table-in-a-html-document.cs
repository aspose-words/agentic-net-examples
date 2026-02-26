using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load an existing HTML document that contains at least one table.
        Document doc = new Document("input.html");

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply desired style options to the table.
        // Example: apply first row formatting and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Optionally, set a built‑in style identifier to see the effect of the options.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Save the modified document.
        doc.Save("output.docx");
    }
}
