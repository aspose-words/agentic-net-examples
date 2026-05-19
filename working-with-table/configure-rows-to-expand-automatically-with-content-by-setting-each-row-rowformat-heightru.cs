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

            // Start a new table.
            Table table = builder.StartTable();

            // Build three rows with two cells each.
            for (int i = 1; i <= 3; i++)
            {
                // First cell with long text to demonstrate automatic row expansion.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1. This is a long piece of text that should cause the row to expand automatically based on its content.");

                // Second cell with short text.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2.");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Ensure every row expands automatically by setting HeightRule to Auto.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.HeightRule = HeightRule.Auto;
            }

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithAutoRows.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
