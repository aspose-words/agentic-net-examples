using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableCellMergeExample
{
    static void Main()
    {
        // Load the HTML document that contains a table.
        Document doc = new Document("input.html");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // -------------------------------------------------
        // Horizontal merge: merge the first two cells of the first row.
        // -------------------------------------------------
        // Set the first cell as the start of the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Merged horizontally"));

        // Set the second cell to merge with the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No content needed for the merged‑into cell.

        // -------------------------------------------------
        // Vertical merge: merge the first two cells of the first column.
        // -------------------------------------------------
        // Set the top cell as the start of the vertical merge.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;
        topCell.FirstParagraph.AppendChild(new Run(doc, "Merged vertically"));

        // Set the cell directly below to merge with the previous cell.
        Cell bottomCell = table.Rows[1].Cells[0];
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;
        // No content needed for the merged‑into cell.

        // Save the modified document back to HTML (or any other format you need).
        doc.Save("output.html");
    }
}
