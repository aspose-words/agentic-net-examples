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
        builder.Writeln("Row 1, Col 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Col 2");
        builder.InsertCell();
        builder.Writeln("Row 1, Col 3");
        builder.EndRow();

        // Third row
        builder.InsertCell();
        builder.Writeln("Row 2, Col 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Col 2");
        builder.InsertCell();
        builder.Writeln("Row 2, Col 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Index of the column to delete (second column, zero‑based index 1).
        int columnIndexToRemove = 1;

        // Remove the cell at the specified column index from each row.
        foreach (Row row in table.Rows)
        {
            // Ensure the row has enough cells.
            if (row.Cells.Count > columnIndexToRemove)
            {
                // Remove the cell; this effectively deletes the column.
                row.Cells[columnIndexToRemove].Remove();
            }
        }

        // Save the modified document.
        string outputPath = "DeletedColumn.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
