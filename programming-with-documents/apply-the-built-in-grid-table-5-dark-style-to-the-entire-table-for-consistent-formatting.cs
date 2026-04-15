using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a few rows with sample data.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("30");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("45");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
            table.StyleIdentifier = StyleIdentifier.GridTable5Dark;
            table.StyleOptions = TableStyleOptions.Default; // Apply style to all table parts.

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "GridTable5Dark.docx");
            doc.Save(outputPath);
        }
    }
}
