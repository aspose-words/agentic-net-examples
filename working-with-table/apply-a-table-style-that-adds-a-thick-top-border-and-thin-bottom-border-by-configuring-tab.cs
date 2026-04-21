using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace TableBorderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑column table with one row.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Apply a thick top border (3 points) and a thin bottom border (0.5 points).
            table.SetBorder(BorderType.Top, LineStyle.Single, 3.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 0.5, Color.Black, true);

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBorders.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // The program ends automatically.
        }
    }
}
