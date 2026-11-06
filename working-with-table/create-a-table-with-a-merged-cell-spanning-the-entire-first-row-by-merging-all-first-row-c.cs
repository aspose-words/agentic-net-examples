using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the table.
        builder.StartTable();

        // ---------- First row (merged cell) ----------
        // Insert first cell and mark it as the start of a horizontal merge.
        builder.InsertCell();
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Header");

        // Insert second cell and mark it as a continuation of the merge.
        builder.InsertCell();
        Cell secondCell = (Cell)builder.CurrentParagraph.ParentNode;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged continuation cells.

        // Insert third cell and also mark it as a continuation of the merge.
        builder.InsertCell();
        Cell thirdCell = (Cell)builder.CurrentParagraph.ParentNode;
        thirdCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // ---------- Second row (regular cells) ----------
        builder.InsertCell();
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.Write("Cell 2");

        builder.InsertCell();
        builder.Write("Cell 3");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "MergedCellTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
