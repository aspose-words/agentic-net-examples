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

        // Begin a table.
        Table table = builder.StartTable();

        int rowCount = 6;   // Total number of rows.
        int colCount = 3;   // Number of cells per row.

        for (int row = 0; row < rowCount; row++)
        {
            // Apply a light gray shading to every second row (odd index).
            if (row % 2 == 1)
                builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            else
                builder.CellFormat.Shading.ClearFormatting(); // No shading for other rows.

            // Populate the cells of the current row.
            for (int col = 0; col < colCount; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }

            // Finish the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the working directory.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Output file was not created: {outputPath}");
    }
}
