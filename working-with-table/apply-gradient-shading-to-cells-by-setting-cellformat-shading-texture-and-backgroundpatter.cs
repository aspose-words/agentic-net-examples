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

        // Start a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        // Apply a diagonal cross texture with a light blue background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Writeln("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        // Apply a horizontal texture with a light green background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Writeln("Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        // Apply a vertical texture with a light coral background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureVertical;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.Writeln("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        // Apply a cross texture with a light yellow background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        builder.Writeln("Cell 4");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "GradientShadingTable.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
