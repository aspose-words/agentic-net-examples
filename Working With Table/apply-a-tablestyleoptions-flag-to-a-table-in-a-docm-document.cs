using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load an existing DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Retrieve the first table in the document (if any).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply desired style options to the table.
        // Example: apply first row formatting and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document as DOCM.
        doc.Save("OutputDocument.docm");
    }
}
