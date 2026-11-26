using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Create 3 rows and 4 columns.
        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 4; col++)
            {
                // Insert a cell and add some text.
                builder.InsertCell();
                builder.Writeln($"Row {row}, Cell {col}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Define output file path.
        string outputPath = "TableExample.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
