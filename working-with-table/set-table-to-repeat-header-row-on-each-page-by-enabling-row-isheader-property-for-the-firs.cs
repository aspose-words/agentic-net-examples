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

        // Start a new table.
        Table table = builder.StartTable();

        // ----- Header row (repeated on each page) -----
        builder.RowFormat.HeadingFormat = true; // Mark this row as a heading.
        builder.InsertCell();
        builder.Write("Header Column 1");
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.EndRow();

        // ----- Data rows (regular rows) -----
        builder.RowFormat.HeadingFormat = false; // Subsequent rows are not headings.

        // Add enough rows to make the table span multiple pages.
        for (int i = 1; i <= 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Column 1");
            builder.InsertCell();
            builder.Write($"Row {i}, Column 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Define output path and ensure the directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWithRepeatingHeader.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
