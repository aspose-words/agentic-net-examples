using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is taken from the first row's cell count.
            // If the table has no rows, column count is zero.
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table found: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally, save the document (e.g., after inspection or modifications).
        doc.Save("OutputDocument.docx");
    }
}
