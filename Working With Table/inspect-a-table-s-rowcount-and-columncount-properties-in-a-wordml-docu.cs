using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the WORDML (or any supported) document.
        // Replace the path with the actual location of your document.
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row nodes in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row.
            // Guard against an empty table (should not happen for a valid table).
            int columnCount = 0;
            if (rowCount > 0 && table.FirstRow != null)
                columnCount = table.FirstRow.Cells.Count;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally, save the document if any modifications were made.
        // doc.Save(@"C:\Docs\Output.docx");
    }
}
