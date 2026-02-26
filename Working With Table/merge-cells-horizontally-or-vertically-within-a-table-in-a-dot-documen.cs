using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // Create a new document and build a table with merged cells.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // ---------- First row (horizontal merge) ----------
        // First cell – marks the start of a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merged with the cell to its left.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");

        builder.EndRow();

        // ---------- Second row (vertical merge) ----------
        // First cell – marks the start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // Second cell – merged with the cell above it.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Third cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the newly created document.
        doc.Save("MergedCells.docx");

        // -------------------------------------------------
        // Load the saved document and apply additional merges.
        // -------------------------------------------------
        Document loadedDoc = new Document("MergedCells.docx");
        Table table = loadedDoc.FirstSection.Body.Tables[0];

        // Merge the first two cells of the first row horizontally.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Merge the first cell of the first column vertically across two rows.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;
        Cell bottomCell = table.Rows[1].Cells[0];
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document.
        loadedDoc.Save("MergedCells_Modified.docx");
    }
}
