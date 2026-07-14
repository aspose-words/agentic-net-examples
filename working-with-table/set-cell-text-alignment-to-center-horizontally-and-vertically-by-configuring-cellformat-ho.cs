using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAlignment
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

            // First cell: set vertical alignment via CellFormat,
            // and horizontal alignment via ParagraphFormat (center).
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Centered");

            // Second cell: same alignment settings.
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Centered");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TableAlignment.docx");
            doc.Save(outputPath);
        }
    }
}
