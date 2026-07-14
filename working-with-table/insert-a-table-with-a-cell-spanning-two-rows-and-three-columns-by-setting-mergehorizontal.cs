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
        // Cell (0,0) – first cell of the merged block (starts both horizontal and vertical merge).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // Start horizontal merge.
        builder.CellFormat.VerticalMerge = CellMerge.First;   // Start vertical merge.
        builder.Write("Merged cell spanning 2 rows and 3 columns.");

        // Cell (0,1) – continues horizontal merge, starts its own vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge.
        builder.CellFormat.VerticalMerge = CellMerge.First;      // Start vertical merge for this column.

        // Cell (0,2) – continues horizontal merge, starts its own vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge.
        builder.CellFormat.VerticalMerge = CellMerge.First;      // Start vertical merge for this column.

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cell (1,0) – continuation of both horizontal and vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge.
        builder.CellFormat.VerticalMerge = CellMerge.Previous;   // Continue vertical merge.

        // Cell (1,1) – continuation of both merges for the second column.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Cell (1,2) – continuation of both merges for the third column.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        doc.Save("MergedTable.docx");
    }
}
