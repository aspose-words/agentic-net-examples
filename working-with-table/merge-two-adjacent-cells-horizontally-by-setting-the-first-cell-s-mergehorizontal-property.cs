using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMerge
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // ---- First row: two cells merged horizontally ----
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged cells");

            // Insert the second cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text is needed for the merged cell.

            // End the first row.
            builder.EndRow();

            // ---- Second row: normal cells (for visual comparison) ----
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Cell 1");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Cell 2");

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            const string outputPath = "MergedCells.docx";
            doc.Save(outputPath);
        }
    }
}
