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
        int rowCount = 6;
        int columnCount = 3;

        // Build the table.
        builder.StartTable();
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < columnCount; c++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {r + 1}, Cell {c + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply shading to every second row (index 1,3,5,...).
        for (int i = 0; i < table.Rows.Count; i++)
        {
            if (i % 2 == 1) // Every second row (1‑based even rows)
            {
                Row row = table.Rows[i];
                // Apply shading to each cell in the row.
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.ForegroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
