using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing Word document that contains at least one table.
        Document doc = new Document("Input.docx");

        // Access the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // ---------- Horizontal merge ----------
        // Merge the first two cells of the first row.
        // The leftmost cell becomes the start of the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // The cell to the right merges into the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---------- Vertical merge ----------
        // Merge the first cell of the first two rows.
        // The top cell starts the vertical merge.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        // The cell directly below merges into the previous (top) cell.
        Cell belowCell = table.Rows[1].Cells[0];
        belowCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
