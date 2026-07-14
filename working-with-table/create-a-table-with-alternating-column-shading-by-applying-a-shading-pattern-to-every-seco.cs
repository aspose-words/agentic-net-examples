using System;
using System.Drawing;
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

        // Define table dimensions.
        int rows = 5;
        int columns = 4;

        // Start the table.
        Table table = builder.StartTable();

        // Build the table row by row.
        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < columns; col++)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply shading to every second column (1‑based index: columns 2,4,...).
                if ((col + 1) % 2 == 0)
                {
                    builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
                else
                {
                    // Ensure no shading on other columns.
                    builder.CellFormat.Shading.ClearFormatting();
                }

                // Write some sample text.
                builder.Write($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnShading.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
