using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the RTF document (replace with your actual file path)
        Document doc = new Document(@"C:\Path\To\InputDocument.rtf");

        // Get the collection of tables in the first section's body
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Number of rows in the table
            int rowCount = table.Rows.Count;

            // Number of columns is determined by the cell count of the first row (if any)
            int columnCount = 0;
            if (table.FirstRow != null)
                columnCount = table.FirstRow.Cells.Count;

            Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document (unchanged) to a new file
        doc.Save(@"C:\Path\To\OutputDocument.rtf");
    }
}
