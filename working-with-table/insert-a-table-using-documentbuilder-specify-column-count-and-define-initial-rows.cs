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

        // Define the number of columns and rows.
        int columnCount = 3;
        int rowCount = 2;

        // Start the table.
        builder.StartTable();

        // Build the table rows and cells.
        for (int row = 1; row <= rowCount; row++)
        {
            for (int col = 1; col <= columnCount; col++)
            {
                // Insert a new cell and write some text into it.
                builder.InsertCell();
                builder.Write($"Row {row}, Col {col}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
