using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // For TextureIndex

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define table dimensions.
        int rows = 10;
        int cols = 3;

        // Build the table.
        builder.StartTable();
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {r + 1}, Cell {c + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Apply alternating shading based on row index parity.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];
        for (int i = 0; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            // Determine the background color for this row.
            Color bgColor = (i % 2 == 0) ? Color.LightGray : Color.White;

            // Apply shading to each cell in the row.
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid; // Correct enum member
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
