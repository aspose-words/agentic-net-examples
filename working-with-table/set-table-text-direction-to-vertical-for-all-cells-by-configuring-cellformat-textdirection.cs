using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableVerticalTextExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Configure the builder to apply vertical text orientation to every cell that will be created.
            builder.CellFormat.Orientation = TextOrientation.Upward; // Upward makes text run vertically.

            // Build a simple 3x2 table.
            Table table = builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();

            // Third row
            builder.InsertCell();
            builder.Write("Cell 1, Row 3");
            builder.InsertCell();
            builder.Write("Cell 2, Row 3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to a file in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalText.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // The program ends here without waiting for user input.
        }
    }
}
