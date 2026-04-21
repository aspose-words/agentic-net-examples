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

        // Start building a table.
        Table table = builder.StartTable();

        // ----- First row -----
        // First cell: start a vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically");

        // Second cell: regular (no merge).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // ----- Second row -----
        // First cell: merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text is written to a merged cell.

        // Second cell: regular (no merge).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Another normal cell");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the resulting document.
        doc.Save("MergedCells.docx");
    }
}
