using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the WORDML (or DOCX) document.
        Document doc = new Document("Input.docx");

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any rows exist).
            int columnCount = 0;
            if (rowCount > 0)
                columnCount = table.Rows[0].Cells.Count;

            Console.WriteLine($"Table found: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document after inspection.
        doc.Save("Output.docx");
    }
}
