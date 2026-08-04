using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableUnmergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a sample document with a table that contains horizontally merged cells.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            Table table = builder.StartTable();

            // First row: three cells merged horizontally.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Content");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Add a normal cell after the merged group.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Normal Cell");

            builder.EndRow();

            // Second row: regular cells (no merging) for comparison.
            builder.InsertCell();
            builder.Write("Row2 Cell1");
            builder.InsertCell();
            builder.Write("Row2 Cell2");
            builder.InsertCell();
            builder.Write("Row2 Cell3");
            builder.InsertCell();
            builder.Write("Row2 Cell4");
            builder.EndRow();

            builder.EndTable();

            // Save the original document (optional, just for reference).
            string originalPath = Path.Combine(Environment.CurrentDirectory, "OriginalTable.docx");
            doc.Save(originalPath);

            // Ensure that merged cells are represented by merge flags.
            // This converts any width‑based merges to HorizontalMerge flags.
            table.ConvertToHorizontallyMergedCells();

            // Process each row to split merged cells into separate cells.
            foreach (Row row in table.Rows)
            {
                // Use a copy of the cell collection because we will modify it during iteration.
                Cell[] cells = row.Cells.ToArray();

                for (int i = 0; i < cells.Length; i++)
                {
                    Cell cell = cells[i];
                    if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                    {
                        // Determine how many cells are part of this merged group.
                        int mergeCount = 1;
                        int j = i + 1;
                        while (j < cells.Length && cells[j].CellFormat.HorizontalMerge == CellMerge.Previous)
                        {
                            mergeCount++;
                            j++;
                        }

                        // For each additional cell in the merged group, insert a new cell with the same content.
                        for (int k = 1; k < mergeCount; k++)
                        {
                            // Clone the original cell (deep clone) to copy its paragraphs.
                            Cell newCell = (Cell)cell.Clone(true);
                            // Ensure the new cell is not marked as merged.
                            newCell.CellFormat.HorizontalMerge = CellMerge.None;
                            // Insert the new cell after the original cell (or after the previously inserted one).
                            row.InsertAfter(newCell, cell);
                            // Update the reference cell so subsequent inserts are placed correctly.
                            cell = newCell;
                        }

                        // After splitting, clear the merge flag on the original first cell.
                        cells[i].CellFormat.HorizontalMerge = CellMerge.None;
                    }
                }

                // Finally, clear any remaining merge flags in the row (e.g., cells that were only 'Previous').
                foreach (Cell c in row.Cells)
                {
                    if (c.CellFormat.HorizontalMerge != CellMerge.None)
                        c.CellFormat.HorizontalMerge = CellMerge.None;
                }
            }

            // Save the transformed document.
            string resultPath = Path.Combine(Environment.CurrentDirectory, "UnmergedTable.docx");
            doc.Save(resultPath);
        }
    }
}
