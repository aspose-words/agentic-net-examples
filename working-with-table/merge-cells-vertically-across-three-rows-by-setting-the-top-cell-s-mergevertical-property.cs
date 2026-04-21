using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ---------- Row 1 ----------
        // First cell: start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically across three rows.");

        // Second cell: regular content.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // ---------- Row 2 ----------
        // First cell: merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // Second cell: regular content.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // ---------- Row 3 ----------
        // First cell: continue merging.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // Second cell: regular content.
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCellsVertical.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Load the document again to verify the merge flag on the top cell.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        Cell topCell = loadedTable.Rows[0].Cells[0];
        if (topCell.CellFormat.VerticalMerge != CellMerge.First)
            throw new Exception("Vertical merge flag not set correctly on the top cell.");
    }
}
