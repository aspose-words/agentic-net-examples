using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeExample
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

            // ---- First row: a single cell that spans the full width ----
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("This cell spans the full width of the table.");

            // Insert additional cells in the same row and merge them with the first cell.
            // The number of cells determines how many columns the table would have.
            int additionalCells = 2; // Adjust as needed for the desired width.
            for (int i = 0; i < additionalCells; i++)
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            }

            // End the first row.
            builder.EndRow();

            // Reset merge settings for any subsequent rows.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // ---- Optional second row: normal cells to demonstrate the table structure ----
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Indicate successful completion (no interactive prompts).
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
