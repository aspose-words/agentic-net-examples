using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTM template document.
        Document doc = new Document("Template.dotm");

        // Retrieve all tables in the document (including those inside other nodes).
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and output its row and column counts.
        foreach (Table table in tableNodes)
        {
            // Number of rows in the current table.
            int rowCount = table.Rows.Count;

            // Number of columns is determined by the cell count of the first row (if any rows exist).
            int columnCount = 0;
            if (rowCount > 0)
                columnCount = table.Rows[0].Cells.Count;

            Console.WriteLine($"Table found: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (optional – demonstrates the required save lifecycle step).
        doc.Save("Output.docx");
    }
}
