using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableCellMergeExample
{
    static void Main()
    {
        // Load an existing PDF document that contains a table.
        Document doc = new Document("Input.pdf");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // ----- Horizontal merge (first row, first two cells) -----
        // Mark the leftmost cell as the start of the merged range.
        table.Rows[0].Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
        // Mark the adjacent cell to merge it with the previous cell.
        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;

        // ----- Vertical merge (first column, first two rows) -----
        // Mark the top cell as the start of the vertically merged range.
        table.Rows[0].Cells[0].CellFormat.VerticalMerge = CellMerge.First;
        // Mark the cell below to merge it with the previous cell vertically.
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document back to PDF.
        doc.Save("Output.pdf");
    }
}
