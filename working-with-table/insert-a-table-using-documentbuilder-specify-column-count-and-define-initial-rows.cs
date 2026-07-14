using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the number of columns and rows for the table.
            int columnCount = 3;
            int rowCount = 2;

            // Start the table.
            Table table = builder.StartTable();

            // Populate the table with the specified rows and columns.
            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {row + 1}, Cell {col + 1}");
                }

                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to a file in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableExample.docx");
            doc.Save(outputPath);

            // Verify that the file was created successfully.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // Indicate successful completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
