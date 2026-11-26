using System;
using System.IO;
using System.Drawing;
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
            builder.StartTable();

            // First cell with gradient shading.
            builder.InsertCell();
            builder.Writeln("Gradient Cell 1");
            Cell cell1 = (Cell)builder.CurrentParagraph.ParentNode;
            cell1.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalDown;
            cell1.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            cell1.CellFormat.Shading.ForegroundPatternColor = Color.White;

            // Second cell with a different gradient shading.
            builder.InsertCell();
            builder.Writeln("Gradient Cell 2");
            Cell cell2 = (Cell)builder.CurrentParagraph.ParentNode;
            cell2.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalUp;
            cell2.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
            cell2.CellFormat.Shading.ForegroundPatternColor = Color.Yellow;

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            string outputPath = "GradientShadingTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
