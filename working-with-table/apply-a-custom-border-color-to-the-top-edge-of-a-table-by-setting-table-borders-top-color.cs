using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBorderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a simple 2x2 grid.
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

            // Finish the table.
            builder.EndTable();

            // Apply a custom color to the top border of the table.
            // Use SetBorder because Table.Borders is not guaranteed to exist.
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Crimson, true);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TableWithTopBorder.docx");
            doc.Save(outputPath);
        }
    }
}
