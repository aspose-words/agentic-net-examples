using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with two cells.
        Table table = builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the row.
        builder.EndRow();

        // -------------------------------------------------
        // Merge the two cells horizontally.
        // Aspose.Words merges cells by setting the HorizontalMerge property.
        // The first cell gets CellMerge.First, the following cell gets CellMerge.Previous.
        // -------------------------------------------------
        Cell firstCell = table.FirstRow.Cells[0];
        Cell secondCell = table.FirstRow.Cells[1];

        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // -------------------------------------------------
        // Split the merged cell back into two separate cells.
        // Splitting is achieved by resetting the HorizontalMerge property to None.
        // -------------------------------------------------
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;

        // Save the resulting document.
        string outputPath = "SplitCellExample.docx";
        doc.Save(outputPath);
    }
}
