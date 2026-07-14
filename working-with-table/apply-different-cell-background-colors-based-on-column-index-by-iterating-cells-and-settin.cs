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

        // Build a sample table with 3 rows and 4 columns.
        Table table = builder.StartTable();

        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply different background colors to cells based on their column index.
        // Column indices are zero‑based within each row.
        Color[] columnColors = new Color[]
        {
            Color.LightCoral,   // Column 0
            Color.LightGreen,   // Column 1
            Color.LightBlue,    // Column 2
            Color.LightYellow   // Column 3
        };

        foreach (Row row in table.Rows)
        {
            for (int i = 0; i < row.Cells.Count; i++)
            {
                Cell cell = row.Cells[i];
                // Guard against out‑of‑range if the table has more columns than colors defined.
                Color bgColor = columnColors[i % columnColors.Length];
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableCellShading.docx");
        doc.Save(outputPath);
    }
}
