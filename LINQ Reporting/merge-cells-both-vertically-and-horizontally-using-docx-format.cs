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

        // Start a table that will contain a 2x2 merged cell block.
        builder.StartTable();

        // ---- First row ----

        // Cell (0,0): first cell of the merged block (both horizontal and vertical).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged 2x2");

        // Cell (0,1): continues horizontal merge, starts vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        // No content needed for merged cells.

        // End the first row.
        builder.EndRow();

        // ---- Second row ----

        // Cell (1,0): starts horizontal merge, continues vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No content needed.

        // Cell (1,1): continues both horizontal and vertical merges.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No content needed.

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document in DOCX format.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);
    }
}
