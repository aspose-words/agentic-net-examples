using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with 3 columns and 3 rows.
        builder.StartTable();

        // First row (header)
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Writeln("Row1 Col1");
        builder.InsertCell();
        builder.Writeln("Row1 Col2");
        builder.InsertCell();
        builder.Writeln("Row1 Col3");
        builder.EndRow();

        // Third row
        builder.InsertCell();
        builder.Writeln("Row2 Col1");
        builder.InsertCell();
        builder.Writeln("Row2 Col2");
        builder.InsertCell();
        builder.Writeln("Row2 Col3");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Delete the second column (index 1) by removing the cell at that index from each row.
        int columnIndexToRemove = 1;
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > columnIndexToRemove)
        {
            foreach (Row row in table.Rows)
            {
                // Ensure the row has enough cells before attempting removal.
                if (row.Cells.Count > columnIndexToRemove)
                    row.Cells.RemoveAt(columnIndexToRemove);
                else
                    throw new InvalidOperationException("A row does not contain enough cells to remove the specified column.");
            }
        }
        else
        {
            throw new InvalidOperationException("The table does not have enough columns to remove.");
        }

        // Save the modified document.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
