using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        // ---------- First Row ----------
        // Cell (1,1) – first cell of the merged block (spans 2 rows x 3 columns).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged cell (2 rows x 3 cols)");

        // Cells (1,2) and (1,3) – continue horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No vertical merge needed for these cells in the first row.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Reset merge settings before adding a normal cell.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // Cell (1,4) – regular cell.
        builder.InsertCell();
        builder.Write("R1C4");

        // End first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cell (2,1) – continuation of vertical merge, first in horizontal range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cells.
        // Cell (2,2) – continuation both horizontally and vertically.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // Cell (2,3) – same as above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Reset merge settings before the next normal cell.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // Cell (2,4) – regular cell.
        builder.InsertCell();
        builder.Write("R2C4");

        // End second row.
        builder.EndRow();

        // ---------- Third Row (no merged cells) ----------
        builder.InsertCell();
        builder.Write("R3C1");
        builder.InsertCell();
        builder.Write("R3C2");
        builder.InsertCell();
        builder.Write("R3C3");
        builder.InsertCell();
        builder.Write("R3C4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        doc.Save("MergedTable.docx");
    }
}
