using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Retrieve the first table in the document (adjust the index if needed).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply desired TableStyleOptions flags to the table.
        // Example: apply formatting to the first row and enable row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document back to DOCM format.
        doc.Save("Output.docm");
    }
}
