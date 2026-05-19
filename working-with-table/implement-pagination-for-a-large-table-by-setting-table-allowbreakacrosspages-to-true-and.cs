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

            // Start a table.
            Table table = builder.StartTable();

            // Define number of rows and columns for the large table.
            const int rowCount = 100;
            const int columnCount = 3;

            // Build the table rows.
            for (int i = 1; i <= rowCount; i++)
            {
                // Insert cells for the current row.
                for (int j = 1; j <= columnCount; j++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {i}, Column {j}");
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Enable pagination across pages for each row and set a uniform height.
            foreach (Row row in table.Rows)
            {
                // Allow the row to break across pages.
                row.RowFormat.AllowBreakAcrossPages = true;

                // Set row height to 20 points with the AtLeast rule so it can grow if needed.
                row.RowFormat.Height = 20;
                row.RowFormat.HeightRule = HeightRule.AtLeast;
            }

            // Save the document to a file.
            string outputPath = "LargeTablePagination.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
