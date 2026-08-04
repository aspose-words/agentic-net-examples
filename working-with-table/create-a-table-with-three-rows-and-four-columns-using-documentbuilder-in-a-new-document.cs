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

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
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

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "Table.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
