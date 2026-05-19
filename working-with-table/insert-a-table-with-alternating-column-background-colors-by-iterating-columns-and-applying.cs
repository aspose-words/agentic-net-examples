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

        // Define table dimensions.
        int rows = 4;
        int columns = 5;

        // Colors for alternating columns.
        Color evenColumnColor = Color.LightGray;
        Color oddColumnColor = Color.White;

        // Start building the table.
        Table table = builder.StartTable();

        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < columns; col++)
            {
                // Insert a new cell in the current row.
                builder.InsertCell();

                // Apply background shading based on column index.
                builder.CellFormat.Shading.BackgroundPatternColor = (col % 2 == 0) ? evenColumnColor : oddColumnColor;

                // Add some sample text to the cell.
                builder.Write($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document to a file.
        string outputPath = Path.Combine(outputDir, "AlternatingColumnsTable.docx");
        doc.Save(outputPath);
    }
}
