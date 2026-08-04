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

        // Start a table with three rows and two columns.
        Table table = builder.StartTable();

        // ---------- Row 1 ----------
        // First cell (column 1) – this will be the top cell of the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First; // Mark as the first merged cell.
        builder.Write("Vertically merged cell (rows 1‑3)");

        // Second cell (column 2) – normal, unmerged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // ---------- Row 2 ----------
        // First cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue vertical merge.
        // No text needed for merged cells; they must be empty.
        // Second cell – independent content.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // ---------- Row 3 ----------
        // First cell – merge with the cell above (third part of the merge).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // Second cell – independent content.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);
    }
}
