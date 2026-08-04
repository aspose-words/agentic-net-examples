using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMarginExample
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

            // Insert a few cells with sample text.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Apply custom margins: left indent and right distance.
            table.LeftIndent = 30;          // Left margin in points.
            table.DistanceRight = 30;       // Right margin in points (alternative to RightIndent).

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TableWithMargins.docx");
            doc.Save(outputPath);
        }
    }
}
