using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template that already contains a table.
        Document doc = new Document("Template.dotx");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // -------------------------------------------------
        // Horizontal merge: merge the first two cells of the first row.
        // -------------------------------------------------
        // Set the first cell as the start of the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        firstCell.CellFormat.VerticalMerge = CellMerge.None; // Ensure vertical merge is not set.
        firstCell.Paragraphs[0].AppendChild(new Run(doc, "Merged horizontally"));

        // Set the second cell to merge with the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        secondCell.CellFormat.VerticalMerge = CellMerge.None;

        // -------------------------------------------------
        // Vertical merge: merge the first cell of the first two rows.
        // -------------------------------------------------
        // Set the top cell as the start of the vertical merge.
        Cell topCell = table.Rows[0].Cells[2];
        topCell.CellFormat.VerticalMerge = CellMerge.First;
        topCell.CellFormat.HorizontalMerge = CellMerge.None;
        topCell.Paragraphs[0].AppendChild(new Run(doc, "Merged vertically"));

        // Set the cell directly below to merge with the previous (top) cell.
        Cell bottomCell = table.Rows[1].Cells[2];
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;
        bottomCell.CellFormat.HorizontalMerge = CellMerge.None;

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        doc.Save("MergedTable.dotx");
    }
}
