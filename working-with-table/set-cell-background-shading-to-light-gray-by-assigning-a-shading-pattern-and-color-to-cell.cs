using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableShading
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

            // First cell – apply light gray shading with a solid texture.
            builder.InsertCell();
            builder.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            builder.Write("Shaded Cell");

            // Second cell – no shading.
            builder.InsertCell();
            builder.Write("Normal Cell");

            // End the first row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableShading.docx");
            doc.Save(outputPath);
        }
    }
}
