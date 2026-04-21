using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table where the first row has a horizontally merged cell.
            Table table = builder.StartTable();

            // First cell – start of the merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged cell");

            // Second cell – continues the merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for the merged part.
            builder.EndRow();

            // Second row – normal, unmerged cells.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Locate the merged cells in the first row and reset their merge flags.
            Row firstRow = table.FirstRow;
            foreach (Cell cell in firstRow.Cells)
            {
                cell.CellFormat.HorizontalMerge = CellMerge.None;
            }

            // Optional validation: after resetting the merge flags the first row should still contain two cells.
            if (firstRow.Cells.Count != 2)
                throw new InvalidOperationException("The table does not contain the expected number of cells after splitting.");

            // Save the resulting document.
            doc.Save("SplitMergedCell.docx");
        }
    }
}
