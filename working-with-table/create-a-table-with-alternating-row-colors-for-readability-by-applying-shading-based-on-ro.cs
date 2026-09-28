using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        int rows = 6;
        int columns = 4;

        // Fill the table with sample text.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {r + 1}, Cell {c + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply alternating row shading (light gray / white) based on row index.
        for (int i = 0; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            Color shadeColor = (i % 2 == 0) ? Color.LightGray : Color.White;

            foreach (Cell cell in row.Cells)
            {
                // Configure solid fill shading.
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                cell.CellFormat.Shading.ForegroundPatternColor = shadeColor;
                cell.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
            }
        }

        // Save the document to disk.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
