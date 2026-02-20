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
        // Mark the leftmost cell as the start of the merged range.
        table.Rows[0].Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
        // Mark the adjacent cell as merged to the previous cell.
        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;

        // -------------------------------------------------
        // Vertical merge: merge the first two cells of the first column.
        // -------------------------------------------------
        // Mark the top cell as the start of the merged range.
        table.Rows[0].Cells[0].CellFormat.VerticalMerge = CellMerge.First;
        // Mark the cell directly below as merged to the previous cell.
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document.
        doc.Save("output.docx");
    }
}
