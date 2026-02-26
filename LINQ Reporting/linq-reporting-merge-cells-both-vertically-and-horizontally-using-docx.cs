using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeCellsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with 3 rows and 3 columns.
        Table table = builder.StartTable();

        // ---------- Row 1 ----------
        // First cell: will be the first cell of a horizontal merge (spans two columns)
        // and also the first cell of a vertical merge (spans two rows).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Horizontal merge start
        builder.CellFormat.VerticalMerge   = CellMerge.First;   // Vertical merge start
        builder.Write("Top‑Left (merged horizontally & vertically)");

        // Second cell: part of the horizontal merge started above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge
        builder.CellFormat.VerticalMerge   = CellMerge.None;     // No vertical merge for this cell
        builder.Write(" ");

        // Third cell: normal cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 1,3");
        builder.EndRow();

        // ---------- Row 2 ----------
        // First cell: continues the vertical merge started in row 1.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // Continue vertical merge
        builder.Write(" "); // Content is ignored for merged cells

        // Second cell: normal cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 2,2");

        // Third cell: normal cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 2,3");
        builder.EndRow();

        // ---------- Row 3 ----------
        // All cells are normal.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 3,1");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 3,2");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 3,3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document in DOCX format.
        string outputPath = Path.Combine(outputDir, "MergedCells.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
