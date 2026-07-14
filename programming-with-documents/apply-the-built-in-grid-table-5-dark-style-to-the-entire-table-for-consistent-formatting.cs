using System;
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

            // Insert the first cell (required before applying any table formatting).
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
            table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

            // Apply the style to all parts of the table.
            table.StyleOptions = TableStyleOptions.Default;

            // Optional: adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Add a few more rows with sample data.
            for (int i = 1; i <= 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Item {i}");
                builder.InsertCell();
                builder.Write($"{i * 10}");
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to the local file system.
            doc.Save("GridTable5Dark.docx");
        }
    }
}
