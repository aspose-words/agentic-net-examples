using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Path to the DOTX template.
        string templatePath = @"C:\Data\Template.dotx";

        // Load the DOTX document.
        Document doc = new Document(templatePath);

        // Get all tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any rows exist).
            int columnCount = 0;
            if (rowCount > 0 && table.FirstRow != null)
                columnCount = table.FirstRow.Cells.Count;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document after inspection (no changes made).
        string outputPath = @"C:\Data\Result.docx";
        doc.Save(outputPath);
    }
}
