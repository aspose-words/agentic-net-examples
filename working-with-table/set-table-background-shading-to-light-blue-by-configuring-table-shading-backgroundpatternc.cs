using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply light‑blue background shading to every cell in the table.
        ApplyTableShading(table, Color.LightBlue);

        // Save the document.
        string fileName = "TableShading.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The output file was not created.");

        // Reload the document and verify the shading color of the first cell.
        Document loadedDoc = new Document(fileName);
        Table loadedTable = (Table)loadedDoc.GetChildNodes(NodeType.Table, true)[0];
        Cell firstCell = loadedTable.Rows[0].Cells[0];
        if (firstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != Color.LightBlue.ToArgb())
            throw new Exception("The table shading color is not light blue.");

        // Indicate success.
        Console.WriteLine($"Table with light blue shading saved to '{fileName}'.");
    }

    private static void ApplyTableShading(Table table, Color color)
    {
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = color;
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            }
        }
    }
}
