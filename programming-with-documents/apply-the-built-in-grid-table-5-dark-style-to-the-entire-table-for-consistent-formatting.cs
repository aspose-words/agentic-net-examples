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

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Insert a header row.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Insert a few data rows.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cherries");
            builder.InsertCell();
            builder.Write("30");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
            table.StyleIdentifier = StyleIdentifier.GridTable5Dark;
            // Apply all style options (first row, last row, banding, etc.) for consistent formatting.
            table.StyleOptions = TableStyleOptions.Default;

            // Define output path (saved in the current working directory).
            string outputPath = Path.Combine(Environment.CurrentDirectory, "GridTable5Dark.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
