using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PaddedTables.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document containing a couple of tables.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First table with two rows and two columns.
        Table table1 = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Add a paragraph between tables.
        builder.Writeln();

        // Second table with three rows and one column.
        Table table2 = builder.StartTable();
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}");
            builder.EndRow();
        }
        builder.EndTable();

        // Save the source document (required by the rule set).
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document (demonstrates loading from file).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Iterate through all tables and apply uniform cell padding of 5 points.
        // -----------------------------------------------------------------
        // Get all Table nodes in the document.
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            // Iterate through each row.
            foreach (Row row in table.Rows)
            {
                // Iterate through each cell in the row.
                foreach (Cell cell in row.Cells)
                {
                    // Set left, top, right, and bottom padding to 5 points.
                    cell.CellFormat.SetPaddings(5, 5, 5, 5);
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // The program finishes without waiting for user input.
    }
}
