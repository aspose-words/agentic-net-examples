using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Insert first cell and apply light gray solid shading.
            builder.InsertCell();
            builder.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            builder.Write("Cell 1");

            // Insert second cell without special shading.
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Insert a second row (no shading needed).
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableShading.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // Optionally, you could load the document again to verify shading,
            // but this example focuses on creation and saving only.
        }
    }
}
