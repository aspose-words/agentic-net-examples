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

        // Start building a table.
        Table table = builder.StartTable();

        // ---------- First row (merged cell) ----------
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell spanning the entire first row.");

        // Insert additional cells in the same row and mark them as merged to the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Reset merge flags before creating the next row to avoid unintended inheritance.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ---------- Second row (regular cells) ----------
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        builder.InsertCell();
        builder.Write("Row 2, Cell 3");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Reload the document and convert any width‑based merges back to merge flags.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        loadedTable.ConvertToHorizontallyMergedCells(); // Ensure merge flags are restored.

        Row firstRow = loadedTable.FirstRow;

        // Validate that the merge flags are as expected.
        if (firstRow.Cells[0].CellFormat.HorizontalMerge != CellMerge.First ||
            firstRow.Cells[1].CellFormat.HorizontalMerge != CellMerge.Previous ||
            firstRow.Cells[2].CellFormat.HorizontalMerge != CellMerge.Previous)
        {
            throw new Exception("Horizontal merge was not applied as expected.");
        }

        // Indicate successful execution.
        Console.WriteLine("Document created and validated successfully.");
    }
}
