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
        int rows = 5;
        int columns = 4;

        // Start building the table.
        builder.StartTable();

        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < columns; col++)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply alternating background colors per column.
                // Even columns (0‑based) get LightGray, odd columns get White.
                builder.CellFormat.Shading.BackgroundPatternColor = (col % 2 == 0) ? Color.LightGray : Color.White;

                // Write some sample text into the cell.
                builder.Writeln($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "AlternatingColumnsTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
