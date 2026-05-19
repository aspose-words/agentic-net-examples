using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        int rowCount = 6;   // Total number of rows.
        int colCount = 3;   // Total number of columns.

        for (int row = 0; row < rowCount; row++)
        {
            // Apply shading to every second row (odd index when zero‑based).
            if (row % 2 == 1)
            {
                // Set a light gray background for the cells of this row.
                builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }
            else
            {
                // Remove any previous shading.
                builder.CellFormat.Shading.ClearFormatting();
            }

            // Insert cells for the current row.
            for (int col = 0; col < colCount; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Col {col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }
    }
}
