using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace GradientShadingTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // First row - three cells with different gradient-like shading.
            // Cell 1
            builder.InsertCell();
            // Apply a diagonal cross texture and a light blue background.
            builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell 1");

            // Cell 2
            builder.InsertCell();
            // Apply a dark diagonal texture and a light green background.
            builder.CellFormat.Shading.Texture = TextureIndex.TextureDarkDiagonalCross;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("Cell 2");

            // Cell 3
            builder.InsertCell();
            // Apply a horizontal texture and a light coral background.
            builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
            builder.Write("Cell 3");

            // End the first row.
            builder.EndRow();

            // Second row - reuse the same shading settings for demonstration.
            // Cell 4
            builder.InsertCell();
            builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Write("Cell 4");

            // Cell 5
            builder.InsertCell();
            builder.CellFormat.Shading.Texture = TextureIndex.TextureDarkDiagonalCross;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Write("Cell 5");

            // Cell 6
            builder.InsertCell();
            builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
            builder.Write("Cell 6");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "GradientShadingTable.docx");
            doc.Save(outputPath);
        }
    }
}
