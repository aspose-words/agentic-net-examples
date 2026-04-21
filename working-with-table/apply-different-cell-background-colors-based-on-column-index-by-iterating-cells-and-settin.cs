using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3‑row by 4‑column table.
        Table table = builder.StartTable();

        int rows = 3;
        int cols = 4;
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

        // Iterate through each cell and set its background color based on column index.
        for (int r = 0; r < table.Rows.Count; r++)
        {
            Row row = table.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                Cell cell = row.Cells[c];
                Color bgColor = c switch
                {
                    0 => Color.LightBlue,
                    1 => Color.LightGreen,
                    2 => Color.LightYellow,
                    _ => Color.LightPink,
                };
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document to disk.
        string outputPath = "ColoredTable.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
