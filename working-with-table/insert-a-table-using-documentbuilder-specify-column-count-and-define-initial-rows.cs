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
            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");

            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters for the table.
            int columnCount = 3; // Number of columns.
            int rowCount = 2;    // Number of initial rows.

            // Start building the table.
            Table table = builder.StartTable();

            // Populate the table with the specified rows and columns.
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {row}, Cell {col}");
                }

                // End the current row before starting the next one.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the specified path.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

            // Optionally, inform that the process completed successfully.
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
