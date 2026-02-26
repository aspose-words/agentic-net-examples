using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMerge
{
    class Program
    {
        static void Main()
        {
            // Load the existing DOTM template.
            // Replace with the actual path to your .dotm file.
            string inputPath = @"C:\Docs\Template.dotm";
            Document doc = new Document(inputPath);

            // Assume the document contains at least one table.
            Table table = doc.FirstSection.Body.Tables[0];

            // -------------------------------------------------
            // Horizontal merge: merge the first two cells of the first row.
            // -------------------------------------------------
            // The leftmost cell becomes the start of the merged range.
            Cell firstCell = table.Rows[0].Cells[0];
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;

            // The cell to the right is merged into the previous cell.
            Cell secondCell = table.Rows[0].Cells[1];
            secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

            // -------------------------------------------------
            // Vertical merge: merge the first two cells of the first column.
            // -------------------------------------------------
            // The top cell becomes the start of the merged range.
            Cell topCell = table.Rows[0].Cells[0];
            topCell.CellFormat.VerticalMerge = CellMerge.First;

            // The cell directly below is merged into the previous (top) cell.
            Cell bottomCell = table.Rows[1].Cells[0];
            bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;

            // -------------------------------------------------
            // Save the modified document.
            // Replace with the desired output path.
            // -------------------------------------------------
            string outputPath = @"C:\Docs\Result.docx";
            doc.Save(outputPath);
        }
    }
}
