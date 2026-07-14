using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder for constructing content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. The builder will now operate inside this table.
        Table table = builder.StartTable();

        // ---------- First Row ----------
        // First cell: start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically");

        // Second cell: regular, not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 1, Col 2");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // First cell: merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // Second cell: regular.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 2, Col 2");

        builder.EndRow();

        // ---------- Third Row ----------
        // First cell: continue merging vertically.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell: regular.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 3, Col 2");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "VerticalMergeTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");

        // Load the saved document and output the vertical merge state of each cell in the first column.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        foreach (Row row in loadedTable.Rows)
        {
            // Get the row index within the table (zero‑based) and add 1 for display.
            int rowNumber = loadedTable.IndexOf(row) + 1;
            Cell firstCell = row.FirstCell;
            Console.WriteLine($"Cell at row {rowNumber} vertical merge: {firstCell.CellFormat.VerticalMerge}");
        }
    }
}
