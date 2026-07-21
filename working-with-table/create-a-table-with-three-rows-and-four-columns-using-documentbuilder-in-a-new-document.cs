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
        builder.StartTable();

        // Create 3 rows.
        for (int row = 0; row < 3; row++)
        {
            // Create 4 columns (cells) in each row.
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");
        doc.Save(outputPath);
    }
}
