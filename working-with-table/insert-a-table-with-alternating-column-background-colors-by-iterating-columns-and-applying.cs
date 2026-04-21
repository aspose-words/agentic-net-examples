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
        int rowCount = 4;
        int columnCount = 5;

        // Start the table.
        Table table = builder.StartTable();

        // Build the table rows and cells.
        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < columnCount; col++)
            {
                // Choose background color based on column index (alternating).
                Color bgColor = (col % 2 == 0) ? Color.LightGray : Color.White;

                // Apply shading to the upcoming cell.
                builder.CellFormat.Shading.BackgroundPatternColor = bgColor;

                // Insert the cell and write its content.
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Col {col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnShading.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
