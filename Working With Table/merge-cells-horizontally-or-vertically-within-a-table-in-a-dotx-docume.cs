using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source DOTX template and the resulting DOCX file.
            string inputDotxPath = @"C:\Docs\Template.dotx";
            string outputDocxPath = @"C:\Docs\MergedCells.docx";

            MergeCellsInTable(inputDotxPath, outputDocxPath);
        }

        /// <summary>
        /// Loads a DOTX document, merges cells horizontally and vertically in its first table,
        /// and saves the result as a DOCX file.
        /// </summary>
        /// <param name="inputPath">Full path to the source DOTX file.</param>
        /// <param name="outputPath">Full path where the modified document will be saved.</param>
        static void MergeCellsInTable(string inputPath, string outputPath)
        {
            // Load the DOTX template.
            Document doc = new Document(inputPath);

            // Ensure the document contains at least one table.
            Table table = doc.FirstSection?.Body?.Tables?.Count > 0
                ? doc.FirstSection.Body.Tables[0]
                : null;

            if (table == null)
                throw new InvalidOperationException("The document does not contain any tables.");

            // -------------------------------------------------
            // Example 1: Horizontal merge (first row, first two cells)
            // -------------------------------------------------
            // Set the first cell as the start of a horizontal merge range.
            Cell firstCell = table.Rows[0].Cells[0];
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;
            firstCell.FirstParagraph?.AppendChild(new Run(doc, "Horizontally merged cells"));

            // Set the second cell to merge with the previous cell.
            Cell secondCell = table.Rows[0].Cells[1];
            secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed in the merged cell; it will be hidden.

            // -------------------------------------------------
            // Example 2: Vertical merge (first column, first two rows)
            // -------------------------------------------------
            // Set the cell in the first row, first column as the start of a vertical merge.
            Cell topCell = table.Rows[0].Cells[0];
            topCell.CellFormat.VerticalMerge = CellMerge.First;
            // The text is already added above.

            // Set the cell directly below to merge vertically with the cell above.
            Cell bottomCell = table.Rows[1].Cells[0];
            bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;
            // No text needed in the merged cell.

            // -------------------------------------------------
            // Example 3: Additional merges (optional)
            // -------------------------------------------------
            // Horizontal merge across three cells in the second row.
            if (table.Rows.Count > 1 && table.Rows[1].Cells.Count >= 3)
            {
                Cell startCell = table.Rows[1].Cells[0];
                startCell.CellFormat.HorizontalMerge = CellMerge.First;
                startCell.FirstParagraph?.AppendChild(new Run(doc, "Three‑cell horizontal merge"));

                table.Rows[1].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
                table.Rows[1].Cells[2].CellFormat.HorizontalMerge = CellMerge.Previous;
            }

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
