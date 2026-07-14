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

            // Start a table with a single column.
            Table table = builder.StartTable();

            int rowCount = 6; // Number of rows to create.

            for (int i = 1; i <= rowCount; i++)
            {
                // Insert a cell for the current row.
                builder.InsertCell();

                // Determine vertical merge settings:
                // - Even rows start a merged range (First).
                // - Odd rows (except the first) continue the previous merge (Previous).
                // - The very first row has no merging.
                if (i % 2 == 0)
                {
                    builder.CellFormat.VerticalMerge = CellMerge.First;
                }
                else if (i % 2 == 1 && i > 1)
                {
                    builder.CellFormat.VerticalMerge = CellMerge.Previous;
                }
                else
                {
                    builder.CellFormat.VerticalMerge = CellMerge.None;
                }

                // Write some text into the cell.
                builder.Write($"Row {i}");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalMerge.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Inform the user (no interactive input required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
