using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // For VerticalAlignment enum

namespace TableVerticalAlignmentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 1x1 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Sample cell");
            builder.EndRow();
            builder.EndTable();

            // Set the table's vertical alignment on the page to the middle (center).
            // Use RelativeVerticalAlignment for floating tables.
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            // Define output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalAlignment.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved successfully.");
        }
    }
}
