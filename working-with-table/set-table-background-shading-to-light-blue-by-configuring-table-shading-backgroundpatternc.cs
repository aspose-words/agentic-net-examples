using System;
using System.Drawing;
using System.IO;
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

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table and obtain the Table node.
            table = builder.EndTable();

            // Apply a solid light‑blue background shading to the whole table.
            // TextureSolid with an empty foreground color and the desired background color.
            table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue);

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBackgroundShading.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
