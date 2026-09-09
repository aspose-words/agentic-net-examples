using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to simplify inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Insert 5 rows and 3 columns.
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                // Insert a new cell and write sample text.
                builder.InsertCell();
                builder.Write($"Row {row}, Col {col}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Define the output file path (in the current working directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.docx");

        // Save the document to disk.
        doc.Save(outputPath);
    }
}
