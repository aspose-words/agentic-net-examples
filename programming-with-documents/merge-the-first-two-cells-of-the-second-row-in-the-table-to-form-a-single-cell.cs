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

            // Build a simple 2x3 table (2 rows, 3 columns) for demonstration.
            builder.StartTable();

            // First row – three separate cells.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.InsertCell();
            builder.Write("R1C3");
            builder.EndRow();

            // Second row – three cells that will be merged (first two cells).
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell");

            // Insert the second cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text is written to this cell because it is merged.

            // Insert the third cell (remains independent).
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("R2C3");

            // Finish the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to a file.
            doc.Save("MergedTable.docx");
        }
    }
}
