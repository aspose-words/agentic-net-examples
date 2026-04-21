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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Build 3 rows.
        for (int row = 1; row <= 3; row++)
        {
            // Build 4 columns for each row.
            for (int col = 1; col <= 4; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row}, Cell {col}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Define the output file path (in the current directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Optionally, inform that the operation succeeded.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
