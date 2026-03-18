using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with 5 columns and 2 rows.
            Table table = builder.StartTable();

            // First row – will contain merged cells.
            for (int i = 0; i < 5; i++)
                builder.InsertCell();

            builder.EndRow();

            // Second row – normal cells for visual reference.
            for (int i = 0; i < 5; i++)
            {
                builder.InsertCell();
                builder.Write($"R2C{i + 1}");
            }

            builder.EndRow();
            builder.EndTable();

            // Merge cells in the first row from column 2 spanning 3 columns (2,3,4).
            MergeCellsHorizontally(table, rowIndex: 0, startColumn: 1, columnSpan: 3);

            // Save the document.
            doc.Save("MergedCells.docx");
        }

        /// <summary>
        /// Merges a range of cells horizontally in a given row.
        /// </summary>
        /// <param name="table">The table containing the cells.</param>
        /// <param name="rowIndex">Zero‑based index of the row to modify.</param>
        /// <param name="startColumn">Zero‑based index of the first cell in the merge range.</param>
        /// <param name="columnSpan">Number of columns to merge (must be >= 1).</param>
        static void MergeCellsHorizontally(Table table, int rowIndex, int startColumn, int columnSpan)
        {
            if (columnSpan < 1)
                throw new ArgumentException("columnSpan must be at least 1.");

            Row row = table.Rows[rowIndex];

            // Ensure the row has enough cells.
            while (row.Cells.Count < startColumn + columnSpan)
                row.Cells.Add(new Cell(table.Document));

            // First cell in the range gets the 'First' flag.
            Cell firstCell = row.Cells[startColumn];
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;

            // Remaining cells get the 'Previous' flag.
            for (int i = 1; i < columnSpan; i++)
            {
                Cell cell = row.Cells[startColumn + i];
                cell.CellFormat.HorizontalMerge = CellMerge.Previous;
            }
        }
    }
}
