using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace SplitMergedCellExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with three columns.
            Table table = builder.StartTable();

            // First cell – start of a horizontally merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell");

            // Second cell – merged with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for merged cells.

            // Third cell – independent.
            builder.InsertCell();
            builder.Write("Cell 3");

            // End the first row and the table.
            builder.EndRow();
            builder.EndTable();

            // At this point the first two cells are merged.
            // Now split the merged cells by resetting merge properties.
            Table firstTable = doc.FirstSection.Body.Tables[0];
            Row firstRow = firstTable.Rows[0];

            foreach (Cell cell in firstRow.Cells)
            {
                // Reset both horizontal and vertical merge flags.
                cell.CellFormat.HorizontalMerge = CellMerge.None;
                cell.CellFormat.VerticalMerge = CellMerge.None;
            }

            // Validate that all cells are now unmerged.
            foreach (Cell cell in firstRow.Cells)
            {
                if (cell.CellFormat.HorizontalMerge != CellMerge.None ||
                    cell.CellFormat.VerticalMerge != CellMerge.None)
                {
                    throw new InvalidOperationException("Cell merge flags were not cleared correctly.");
                }
            }

            // Save the resulting document.
            doc.Save("SplitMergedCell.docx");
        }
    }
}
