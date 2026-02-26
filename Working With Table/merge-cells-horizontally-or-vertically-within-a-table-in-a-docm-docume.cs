using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Ensure the table has at least 2 rows and 2 columns.
        table.EnsureMinimum();

        // ---------- Horizontal merge (first row, first two cells) ----------
        // Mark the leftmost cell as the start of the merged range.
        Cell leftCell = table.Rows[0].Cells[0];
        leftCell.CellFormat.HorizontalMerge = CellMerge.First;

        // Mark the cell to the right as merged with the previous cell.
        Cell rightCell = table.Rows[0].Cells[1];
        rightCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---------- Vertical merge (first column, first two rows) ----------
        // Mark the top cell as the start of the vertically merged range.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        // Mark the cell directly below as merged with the previous cell.
        Cell bottomCell = table.Rows[1].Cells[0];
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document as a DOCM file.
        doc.Save("Output.docm");
    }
}
