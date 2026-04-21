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

        // First cell – apply a diagonal cross texture with a light blue background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 1");
        builder.InsertCell();

        // Second cell – apply a vertical texture with a light green background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureVertical;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Write("Cell 2");
        builder.InsertCell();

        // Third cell – apply a horizontal texture with a light coral background.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.Write("Cell 3");
        builder.EndRow();

        // Add a second row with different shading for demonstration.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalDown;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        builder.Write("Cell 4");
        builder.InsertCell();

        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalUp;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightPink;
        builder.Write("Cell 5");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GradientShadingTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The document was not saved successfully.");
        }
    }
}
