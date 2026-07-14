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

        // Start the table.
        Table table = builder.StartTable();

        // Build 3 rows.
        for (int row = 1; row <= 3; row++)
        {
            // Build 4 columns for each row.
            for (int col = 1; col <= 4; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row}, Col {col}");
            }

            // End the current row (except after the last row, EndTable will close it).
            if (row < 3)
                builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
