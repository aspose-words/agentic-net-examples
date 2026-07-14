using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a few rows/cells with sample text.
            Table table = builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Third row
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Set the table's preferred width to 50% of the page width (half page).
            table.PreferredWidth = PreferredWidth.FromPercent(50);

            // Ensure auto‑fit is enabled so columns can adjust dynamically.
            // The default value of AllowAutoFit is true, but we set it explicitly for clarity.
            table.AllowAutoFit = true;

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidth.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
