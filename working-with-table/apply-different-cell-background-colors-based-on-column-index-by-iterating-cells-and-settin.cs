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

        // Build a simple 3x4 table.
        Table table = builder.StartTable();

        int rows = 3;
        int columns = 4;

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply different background colors based on column index.
        // Define a set of colors to use for each column.
        Color[] columnColors = new Color[]
        {
            Color.LightBlue,
            Color.LightGreen,
            Color.LightYellow,
            Color.LightCoral
        };

        // Iterate through each cell and set its shading.
        foreach (Row row in table.Rows)
        {
            for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
            {
                Cell cell = row.Cells[colIndex];
                // Choose color based on column index (wrap if more columns than colors).
                Color bgColor = columnColors[colIndex % columnColors.Length];
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document to a file.
        string outputPath = "TableCellShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
