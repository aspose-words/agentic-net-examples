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

            // Start a table.
            Table table = builder.StartTable();

            // Add five rows. Rows with odd index (1,3) will be treated as "empty" (height = 0).
            for (int i = 0; i < 5; i++)
            {
                // First cell of the row.
                builder.InsertCell();

                if (i % 2 == 0) // rows 0,2,4 are non‑empty
                {
                    // Write text for the first cell.
                    builder.Write($"Row {i + 1}, Cell 0");

                    // Second cell of the same row.
                    builder.InsertCell();
                    // Write text for the second cell.
                    builder.Write($"Row {i + 1}, Cell 1");

                    // Set a visible height for the row.
                    builder.RowFormat.Height = 30;
                    builder.RowFormat.HeightRule = HeightRule.Exactly;
                }
                else
                {
                    // Second cell (still required to keep table structure).
                    builder.InsertCell();

                    // Leave the row without content and keep the default height (0).
                    builder.RowFormat.Height = 0;
                    builder.RowFormat.HeightRule = HeightRule.Auto;
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Remove rows whose height is zero (considered empty).
            // Iterate backwards to avoid index issues while removing.
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                Row row = table.Rows[i];
                if (row.RowFormat.Height == 0)
                {
                    row.Remove();
                }
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
