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

        // Build a table with enough rows to potentially span multiple pages.
        Table table = builder.StartTable();

        // Add 30 rows, each with two cells containing sample text.
        for (int i = 1; i <= 30; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Column 1 - Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
            builder.InsertCell();
            builder.Write($"Row {i}, Column 2 - Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.EndRow();
        }

        builder.EndTable();

        // Prevent each row from breaking across pages.
        foreach (Row row in table.Rows)
        {
            // Setting AllowBreakAcrossPages to false keeps the row together.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.RowKeepTogether.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
