using Aspose.Words;
using Aspose.Words.Tables;

public class TableCellMerging
{
    public void MergeCells()
    {
        // Create a new document (the actual creation is handled by the provided create rule)
        Document doc = new Document();               // <-- create rule
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a 2x2 table
        Table table = builder.StartTable();

        // ---------- First Row ----------
        // First cell – start of a horizontal merge range and also the start of a vertical merge range
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Horizontal merge start
        builder.CellFormat.VerticalMerge   = CellMerge.First;   // Vertical merge start
        builder.Write("Merged Cell (H+V)");

        // Second cell – merge horizontally with the first cell
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge to previous cell horizontally
        builder.CellFormat.VerticalMerge   = CellMerge.None;     // No vertical merge for this cell
        // No text needed for merged cell
        builder.EndRow();

        // ---------- Second Row ----------
        // First cell – merge vertically with the cell above
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;     // No horizontal merge
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // Merge to previous cell vertically
        // No text needed for merged cell

        // Second cell – regular unmerged cell
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Regular Cell");
        builder.EndRow();

        builder.EndTable();

        // Save the document (the actual saving is handled by the provided save rule)
        doc.Save("MergedCells.docx");               // <-- save rule
    }
}

public static class Program
{
    public static void Main()
    {
        new TableCellMerging().MergeCells();
    }
}
