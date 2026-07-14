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

            // First row – default spacing.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row – set double line spacing.
            // HeightRule.Auto allows the row to grow automatically,
            // and Height is set to 24 points (approximately double the default 12‑point line height).
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.RowFormat.Height = 24.0;

            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRowSpacing.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
