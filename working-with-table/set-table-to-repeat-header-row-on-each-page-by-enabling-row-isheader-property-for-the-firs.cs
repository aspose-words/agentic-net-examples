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

        // ----- Header row (will repeat on each page) -----
        builder.RowFormat.HeadingFormat = true; // Enable repeat header.
        builder.InsertCell();
        builder.Write("Header Column 1");
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.EndRow();

        // Reset the flag for normal rows.
        builder.RowFormat.HeadingFormat = false;

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

        // Define output file name.
        string outputPath = "HeaderRepeatingTable.docx";

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");

        // Inform that the process completed successfully.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
