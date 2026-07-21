using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableTextDirection
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

            // Build a 2x2 table. Set each cell's orientation to vertical Asian script
            // (top‑to‑bottom) using the CellFormat.Orientation property.
            for (int row = 0; row < 2; row++)
            {
                for (int col = 0; col < 2; col++)
                {
                    builder.InsertCell();
                    builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableTextDirection.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");
        }
    }
}
