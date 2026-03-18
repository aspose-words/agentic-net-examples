using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 4‑column, 5‑row table.
        int rows = 5;
        int cols = 4;

        Table table = builder.StartTable();

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply alternating background colors to columns.
        // Even columns (0‑based) get LightGray, odd columns get White.
        for (int colIndex = 0; colIndex < cols; colIndex++)
        {
            Color bgColor = (colIndex % 2 == 0) ? Color.LightGray : Color.White;

            foreach (Row row in table.Rows)
            {
                // Ensure the cell exists (tables are rectangular in this example).
                Cell cell = row.Cells[colIndex];
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "AlternatingColumns.docx");
        doc.Save(outputPath);
    }
}
