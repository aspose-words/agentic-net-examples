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

        // Start building a table.
        builder.StartTable();

        int rows = 10;   // Number of rows in the table.
        int cols = 3;    // Number of columns per row.

        // Populate the table with sample text.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {r + 1}, Cell {c + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply alternating shading to every second row (odd‑indexed rows).
        Table table = doc.FirstSection.Body.Tables[0];
        for (int i = 0; i < table.Rows.Count; i++)
        {
            if (i % 2 == 1) // Every second row (1‑based: rows 2,4,6,...)
            {
                Row row = table.Rows[i];
                foreach (Cell cell in row.Cells)
                {
                    // Apply solid light gray shading to each cell in the row.
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.ForegroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document to a file.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
