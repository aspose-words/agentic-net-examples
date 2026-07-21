using System;
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

            // Build three rows with two cells each.
            for (int i = 1; i <= 3; i++)
            {
                // First cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");

                // Second cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");

                // Ensure the row will expand automatically to fit its content.
                builder.RowFormat.HeightRule = HeightRule.Auto;

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // As an extra safety check, set the HeightRule of every row to Auto.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.HeightRule = HeightRule.Auto;
            }

            // Save the document to the local file system.
            string outputPath = "RowsAutoHeight.docx";
            doc.Save(outputPath);
        }
    }
}
