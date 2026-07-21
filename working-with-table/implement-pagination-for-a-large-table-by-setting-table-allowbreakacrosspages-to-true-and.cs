using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTablePagination
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Build a large table (e.g., 50 rows, 2 columns).
            for (int i = 1; i <= 50; i++)
            {
                // Set row height to a reasonable value and allow it to grow if needed.
                builder.RowFormat.Height = 20;               // Height in points.
                builder.RowFormat.HeightRule = HeightRule.AtLeast;

                // First cell.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");

                // Second cell.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Enable breaking rows across pages for pagination.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = true;
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LargeTablePagination.docx");
            doc.Save(outputPath);
        }
    }
}
