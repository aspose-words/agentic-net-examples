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

        // Start a table with two columns.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        Row row1 = builder.EndRow();
        // Prevent this row from breaking across pages.
        row1.RowFormat.AllowBreakAcrossPages = false;

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        Row row2 = builder.EndRow();
        row2.RowFormat.AllowBreakAcrossPages = false;

        // Third row.
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        Row row3 = builder.EndRow();
        row3.RowFormat.AllowBreakAcrossPages = false;

        // Finish the table.
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.AllowBreakAcrossPages.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
