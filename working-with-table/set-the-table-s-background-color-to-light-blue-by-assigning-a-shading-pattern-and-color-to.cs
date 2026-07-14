using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a couple of cells with sample text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Apply a solid light‑blue shading to the whole table.
            // TextureSolid indicates a solid fill; the foreground color is not used for this texture.
            table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue);

            // Finish the table.
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableShading.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // The program ends automatically; no user interaction required.
        }
    }
}
