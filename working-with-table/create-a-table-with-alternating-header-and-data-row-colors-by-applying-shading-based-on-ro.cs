using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // Define number of rows (including header) and columns.
        int totalRows = 6; // 1 header + 5 data rows
        int totalCols = 3;

        // Loop through each row.
        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
        {
            // Apply shading based on row index parity.
            // Even index (0) -> header row, use LightGray.
            // Odd index -> LightBlue, even index after header -> White.
            if (rowIndex == 0)
                builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            else if (rowIndex % 2 == 1)
                builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            else
                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;

            // Insert cells for the current row.
            for (int colIndex = 0; colIndex < totalCols; colIndex++)
            {
                builder.InsertCell();
                // Write sample text indicating row and column.
                if (rowIndex == 0)
                    builder.Write($"Header {colIndex + 1}");
                else
                    builder.Write($"Row {rowIndex}, Col {colIndex + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
