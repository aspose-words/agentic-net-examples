using System;
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

            // Start building a table.
            Table table = builder.StartTable();

            // ---------- First row (merged horizontally across three columns) ----------
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged across 3 columns");

            // Insert the second cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Insert the third cell and also merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // Reset merge setting for subsequent cells.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // ---------- Second row (regular, unmerged cells) ----------
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = "MergedCells.docx";
            doc.Save(outputPath);

            // Inform the user that the file has been created.
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
