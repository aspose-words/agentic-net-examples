using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Tables.docx");

        // Get all tables in the first section.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns – use the cell count of the first row (tables are rectangular).
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (no modifications were made, but saving demonstrates the lifecycle).
        doc.Save("Tables_Inspected.docx");
    }
}
