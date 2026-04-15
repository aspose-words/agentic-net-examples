using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitCellExample
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SplitCell.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with two horizontally merged cells.
        builder.StartTable();

        // First cell – start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell");

        // Second cell – merged to the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the cells of the first row.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[0].Cells[1];

        // Split the merged cells back into two separate cells.
        // To "unmerge" set the HorizontalMerge property of both cells to None.
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;

        // Save the resulting document.
        doc.Save(outputPath);
    }
}
