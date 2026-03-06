using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document that contains a table.
        Document doc = new Document("input.html");

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // -------------------------------------------------
        // Horizontal merge: merge the first two cells of the first row.
        // -------------------------------------------------
        // Mark the leftmost cell as the start of the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // Mark the cell to the right as merged with the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // -------------------------------------------------
        // Vertical merge: merge the first two cells of the first column.
        // -------------------------------------------------
        // Mark the top cell as the start of the vertically merged range.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        // Mark the cell directly below as merged with the previous cell.
        Cell belowCell = table.Rows[1].Cells[0];
        belowCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document back to HTML (or any other format you need).
        doc.Save("output.html");
    }
}
