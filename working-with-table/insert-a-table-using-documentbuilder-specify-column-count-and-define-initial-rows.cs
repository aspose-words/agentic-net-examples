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

            // Associate a DocumentBuilder with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table. The builder's cursor is now inside the first cell.
            Table table = builder.StartTable();

            // Define the number of columns.
            int columnCount = 3;

            // Define the number of initial rows.
            int rowCount = 2;

            // Populate the table with the specified rows and columns.
            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    // Insert a cell (or move to the next cell in the current row).
                    builder.InsertCell();

                    // Write some sample text into the cell.
                    builder.Write($"Row {row + 1}, Column {col + 1}");
                }

                // End the current row and start a new one (except after the last row).
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Determine an output path in the current working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

            // Optionally, inform the user that the operation succeeded.
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
