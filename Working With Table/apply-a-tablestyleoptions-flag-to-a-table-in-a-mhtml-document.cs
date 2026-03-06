using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing MHTML document.
        Document doc = new Document("input.mhtml");

        // Retrieve the first table in the document (if any).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table != null)
        {
            // Apply desired style options to the table.
            // Example: apply formatting to the first row and enable row banding.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
        }

        // Save the modified document back to MHTML format.
        doc.Save("output.mhtml");
    }
}
