using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsMergeCellsExample
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

            // ---------- First Row (merged cells) ----------
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("This cell spans three columns.");

            // Insert the second cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Insert the third cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // Reset merge settings for subsequent rows.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // ---------- Second Row (regular cells) ----------
            builder.InsertCell();
            builder.Write("Cell 1");

            builder.InsertCell();
            builder.Write("Cell 2");

            builder.InsertCell();
            builder.Write("Cell 3");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Optional: indicate success (no interactive prompts required).
            Console.WriteLine("Document created successfully at: " + outputPath);
        }
    }
}
