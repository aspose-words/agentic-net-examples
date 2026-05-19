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

        // Define the number of columns and rows for the table.
        int columnCount = 3;
        int rowCount = 2;

        // Start the table.
        Table table = builder.StartTable();

        // Build the table row by row.
        for (int row = 0; row < rowCount; row++)
        {
            // Insert the required number of cells for this row.
            for (int col = 0; col < columnCount; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }

            // End the current row and start a new one (except after the last row).
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
