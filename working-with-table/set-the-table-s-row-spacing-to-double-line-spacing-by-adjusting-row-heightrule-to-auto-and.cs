using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableRowSpacing
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

            // First row with double line spacing.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            // Set row height rule to Auto and height to approx. double line spacing (24 points).
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.RowFormat.Height = 24;
            builder.EndRow();

            // Second row with the same spacing.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.RowFormat.Height = 24;
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableRowSpacing.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Failed to create the output document.");
        }
    }
}
