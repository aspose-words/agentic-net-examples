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

        // Build a simple 2x1 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply light blue solid shading to every cell in the table.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                cell.CellFormat.Shading.ForegroundPatternColor = Color.LightBlue;
            }
        }

        // Save the document.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
