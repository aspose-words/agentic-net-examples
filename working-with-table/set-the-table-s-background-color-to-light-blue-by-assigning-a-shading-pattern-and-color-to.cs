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

            // Start a table and add a single row with two cells.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Apply a solid light‑blue shading to the whole table.
            table.SetShading(TextureIndex.TextureSolid, Color.LightBlue, Color.Empty);

            // Save the document to a file in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableShading.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // Optionally, you could load the document again to verify the shading,
            // but for this simple example the existence check is sufficient.
        }
    }
}
