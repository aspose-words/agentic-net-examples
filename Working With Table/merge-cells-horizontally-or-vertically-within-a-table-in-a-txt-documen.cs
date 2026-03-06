using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // ---------- First row ----------
        // Insert first cell and mark it as the start of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Insert second cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Insert a third cell that is not merged.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Regular cell");

        // End the first row.
        builder.EndRow();

        // ---------- Second row ----------
        // Insert first cell and mark it as the start of a vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // Insert second cell and merge it vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Insert a third cell that is not merged vertically.
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Regular cell 2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a plain‑text file.
        doc.Save("MergedTable.txt", SaveFormat.Text);
    }
}
