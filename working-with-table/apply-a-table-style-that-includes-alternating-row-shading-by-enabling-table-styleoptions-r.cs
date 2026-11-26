using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in style that supports shading.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Apply alternating row shading (row banding) by shading each cell in the target rows.
        for (int i = 0; i < table.Rows.Count; i++)
        {
            // Shade every other row (e.g., even rows, index 1, 3, ...).
            if (i % 2 == 1)
            {
                Row row = table.Rows[i];
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                }
            }
        }

        // Save the document.
        const string fileName = "TableRowBanding.docx";
        doc.Save(fileName);

        // Verify the file was created.
        if (!System.IO.File.Exists(fileName))
            throw new Exception("Failed to create the output document.");

        // Confirmation (non‑interactive).
        Console.WriteLine($"Document saved to {fileName}");
    }
}
