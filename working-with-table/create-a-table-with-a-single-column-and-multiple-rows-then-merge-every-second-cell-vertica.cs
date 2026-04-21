using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with a single column.
            builder.StartTable();

            int totalRows = 8; // Example: 8 rows (will create 4 merged groups).

            for (int i = 0; i < totalRows; i++)
            {
                // Insert a cell for the current row.
                builder.InsertCell();

                // For every first cell of a pair set VerticalMerge to First,
                // for the second cell set it to Previous to merge with the cell above.
                if (i % 2 == 0)
                {
                    builder.CellFormat.VerticalMerge = CellMerge.First;
                }
                else
                {
                    builder.CellFormat.VerticalMerge = CellMerge.Previous;
                }

                // Add some text to identify the row (only the first cell of each pair will display text).
                builder.Write($"Group {(i / 2) + 1}");

                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Optionally inform the user (no interactive pause required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
