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

        // Start a table.
        builder.StartTable();

        // ---------- First row ----------
        // First cell will be the top‑left cell of a 2x2 merged block.
        builder.InsertCell();
        // Mark this cell as the first cell in a horizontally merged range.
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        // Mark this cell as the first cell in a vertically merged range.
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged 2x2");

        // Second cell in the first row – merge horizontally with the first cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No vertical merge for this cell (it will be merged vertically by the cell above).
        builder.CellFormat.VerticalMerge = CellMerge.None;
        // No content needed; the cell is merged.
        builder.EndRow();

        // ---------- Second row ----------
        // First cell in the second row – merge vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No content needed; the cell is merged.
        // Second cell in the second row – merge horizontally with the cell to its left.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        // No content needed; the cell is merged.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document in DOC format.
        string outputPath = "MergedCells.doc";
        doc.Save(outputPath, SaveFormat.Doc);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
