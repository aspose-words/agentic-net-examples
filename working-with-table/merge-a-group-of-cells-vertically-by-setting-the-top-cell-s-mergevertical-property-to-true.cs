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

        // Start a table. The table will have two columns and three rows.
        Table table = builder.StartTable();

        // ---------- First row ----------
        // First cell – this cell will be the top of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First; // Mark as the first merged cell.
        builder.Write("Merged vertically");

        // Second cell – a regular, unmerged cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // ---------- Second row ----------
        // First cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue vertical merge.
        // No text is written to merged cells beyond the first.

        // Second cell – regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // ---------- Third row ----------
        // First cell – continue merging vertically.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell – regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document to a known location.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "MergedCells.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");
    }
}
