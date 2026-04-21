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

            // Start a table with 5 columns.
            builder.StartTable();

            // ---------- First row (regular cells) ----------
            for (int i = 1; i <= 5; i++)
            {
                builder.InsertCell();
                builder.Write($"Header {i}");
            }
            builder.EndRow();

            // ---------- Second row (cells with horizontal merges) ----------
            // Cell 1 – start a merge that spans 2 columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Span 2 columns");

            // Cell 2 – merged with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Reset merge setting before starting a new merge range.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // Cell 3 – start a merge that spans 3 columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Span 3 columns");

            // Cell 4 – merged with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Cell 5 – merged with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCellsTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Inform that the process completed (no interactive input required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
