using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell (spanning two columns).
        builder.StartTable();

        // First cell – mark as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell");

        // Second cell – merge with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.EndRow();
        builder.EndTable();

        // Locate the merged cells (the first row).
        Table table = doc.FirstSection.Body.Tables[0];
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[0].Cells[1];

        // Split the merged cells by clearing the merge flags.
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;

        // Add text to the now separate cells.
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "SplitCellExample.docx");
        doc.Save(outputPath);
    }
}
