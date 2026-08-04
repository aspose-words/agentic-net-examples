using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output directory and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SplitCell.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell spanning two columns.
        builder.StartTable();
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // First cell in the merge range.
        builder.Write("Merged cell");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with the previous cell.
        builder.EndRow();
        builder.EndTable();

        // Locate the table.
        Table table = doc.FirstSection.Body.Tables[0];

        // -----------------------------------------------------------------
        // Split the merged cell back into two separate cells.
        // Aspose.Words does not provide a SplitCell method on Table.
        // To "split" the merged cells we simply clear the merge flags on both
        // cells in the row. The row already contains two Cell objects – the
        // first marked as CellMerge.First and the second as CellMerge.Previous.
        // Resetting the flags restores them to independent cells.
        // -----------------------------------------------------------------
        Row firstRow = table.Rows[0];
        Cell firstCell = firstRow.Cells[0];
        Cell secondCell = firstRow.Cells[1];

        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;

        // Update the text in the now‑separate cells.
        firstCell.FirstParagraph.Runs.Clear();
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));

        secondCell.FirstParagraph.Runs.Clear();
        secondCell.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));

        // Save the resulting document.
        doc.Save(outputPath);
    }
}
