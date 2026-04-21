using System;
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

            // Start a table.
            Table table = builder.StartTable();

            // Build a large table with 50 rows and 2 columns.
            for (int i = 1; i <= 50; i++)
            {
                // First cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");

                // Second cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Enable pagination for each row and set a minimum height.
            foreach (Row row in table.Rows)
            {
                // Allow the row to break across pages.
                row.RowFormat.AllowBreakAcrossPages = true;

                // Set a minimum height for the row (in points).
                row.RowFormat.Height = 20;
                row.RowFormat.HeightRule = HeightRule.AtLeast;
            }

            // Save the document to the local file system.
            string outputPath = "LargeTablePagination.docx";
            doc.Save(outputPath);
        }
    }
}
