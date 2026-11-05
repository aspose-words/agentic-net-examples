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
        builder.StartTable();

        // First cell with gradient shading.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Writeln("Cell 1");

        // Second cell with a different gradient shading.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Writeln("Cell 2");

        // End the first row.
        builder.EndRow();

        // Third cell with gradient shading.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureVertical;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.Writeln("Cell 3");

        // Fourth cell with gradient shading.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGoldenrodYellow;
        builder.Writeln("Cell 4");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "GradientShadingTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
