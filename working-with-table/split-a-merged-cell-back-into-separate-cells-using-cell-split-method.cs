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

        // -------------------------------------------------
        // Build a table with a horizontally merged cell.
        // -------------------------------------------------
        Table table = builder.StartTable();

        // First cell – start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell content");

        // Second cell – merged with the previous one.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Add two normal cells in the second row.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document that contains the merged cell.
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCell.docx");
        doc.Save(mergedPath);

        // -------------------------------------------------
        // Split the previously merged cell back into separate cells.
        // -------------------------------------------------
        // Aspose.Words does not provide a Cell.Split method, so we recreate the row
        // with the desired number of cells while preserving the original content.
        Row firstRow = table.FirstRow;
        Cell mergedCell = firstRow.FirstCell;

        // Preserve the text that was inside the merged cell.
        string mergedText = mergedCell.GetText().Trim();

        // Remove all existing cells from the first row.
        firstRow.Cells.Clear();

        // Create the first new cell and copy the original merged text.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, mergedText));
        firstRow.AppendChild(cell1);

        // Create the second new (empty) cell.
        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        firstRow.AppendChild(cell2);

        // Create the third cell that corresponds to the original second cell (empty).
        Cell cell3 = new Cell(doc);
        cell3.AppendChild(new Paragraph(doc));
        firstRow.AppendChild(cell3);

        // Verify that the split produced the expected number of cells.
        if (firstRow.Cells.Count != 3)
            throw new InvalidOperationException("Cell split did not produce the expected number of cells.");

        // Save the document after splitting.
        string splitPath = Path.Combine(Directory.GetCurrentDirectory(), "SplitMergedCell.docx");
        doc.Save(splitPath);
    }
}
