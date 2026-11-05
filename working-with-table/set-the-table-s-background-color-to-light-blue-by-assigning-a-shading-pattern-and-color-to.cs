using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply solid light‑blue shading to every cell in the table.
        foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
        {
            cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            cell.CellFormat.Shading.ForegroundPatternColor = Color.LightBlue;
        }

        // Save the document.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Reload the document and verify the shading of the first cell.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChildNodes(NodeType.Table, true)[0];
        Cell firstCell = (Cell)loadedTable.GetChildNodes(NodeType.Cell, true)[0];

        if (firstCell.CellFormat.Shading.Texture != TextureIndex.TextureSolid ||
            firstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb() != Color.LightBlue.ToArgb())
        {
            throw new Exception("Table shading was not applied correctly.");
        }
    }
}
