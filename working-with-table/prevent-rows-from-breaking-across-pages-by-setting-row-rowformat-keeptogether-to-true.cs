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

        // Start a table.
        Table table = builder.StartTable();

        // Configure the row format so that rows will not break across pages.
        // The property that controls this behavior is AllowBreakAcrossPages.
        // Setting it to false keeps the entire row together on one page.
        builder.RowFormat.AllowBreakAcrossPages = false;

        // Build a few rows with sample content.
        for (int i = 1; i <= 5; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not saved correctly.");
        }

        // Inform that the process completed successfully.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
