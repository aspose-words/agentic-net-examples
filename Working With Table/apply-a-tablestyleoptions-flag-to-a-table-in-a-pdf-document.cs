using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load the source PDF document.
        Document doc = new Document("input.pdf");

        // Retrieve the first table in the document (if any).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table != null)
        {
            // Apply desired style options to the table.
            // Example: apply first row formatting and row banding.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
        }

        // Save the modified document as PDF.
        doc.Save("output.pdf");
    }
}
