using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableCellMergeExample
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder for building content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Build a table with merged cells (horizontal merge)
        // -------------------------------------------------
        Table table = builder.StartTable();

        // First row – merge first two cells horizontally.
        builder.InsertCell();                                   // Cell (0,0)
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Mark as first merged cell.
        builder.Write("Horizontally merged cells");

        builder.InsertCell();                                   // Cell (0,1)
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with the cell to the left.

        // Add a third cell that is not merged.
        builder.InsertCell();                                   // Cell (0,2)
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // Second row – normal cells (no merging).
        builder.InsertCell();                                   // Cell (1,0)
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();                                   // Cell (1,1)
        builder.Write("Row 2, Cell 2");

        builder.InsertCell();                                   // Cell (1,2)
        builder.Write("Row 2, Cell 3");
        builder.EndRow();

        // -------------------------------------------------
        // Build a table with merged cells (vertical merge)
        // -------------------------------------------------
        Table vTable = builder.StartTable();

        // First column – merge first two cells vertically.
        builder.InsertCell();                                   // Cell (0,0)
        builder.CellFormat.VerticalMerge = CellMerge.First;     // First cell in vertical merge range.
        builder.Write("Vertically merged cells");

        builder.InsertCell();                                   // Cell (0,1)
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row – second cell normal, first cell merges with above.
        builder.InsertCell();                                   // Cell (1,0)
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Merge with cell above.
        // No content needed for merged cell.

        builder.InsertCell();                                   // Cell (1,1)
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // -------------------------------------------------
        // Save the document.
        // -------------------------------------------------
        doc.Save("MergedCellsTable.docx");
    }
}
