using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableInspector
{
    static void Main()
    {
        // Load the existing DOCM document.
        string inputPath = "input.docm";
        Document doc = new Document(inputPath);

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns is taken from the first row (assuming a regular table).
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document (e.g., after inspection or modifications).
        string outputPath = "output.docm";
        doc.Save(outputPath);
    }
}
