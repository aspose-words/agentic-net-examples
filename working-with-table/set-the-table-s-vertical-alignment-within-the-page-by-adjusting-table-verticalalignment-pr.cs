using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // Needed for VerticalAlignment enum

namespace TableVerticalAlignmentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Build a simple 1x1 table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Sample cell");
            builder.EndRow();
            builder.EndTable();

            // Set the table's vertical alignment on the page to the middle.
            // For floating tables use RelativeVerticalAlignment with the drawing.VerticalAlignment enum.
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalAlignment.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
