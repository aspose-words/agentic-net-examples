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

        // Build a table with 2 rows and 3 columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.InsertCell();
        builder.Writeln("Cell 5");
        builder.InsertCell();
        builder.Writeln("Cell 6");
        builder.EndRow();

        builder.EndTable();

        // Merge the first two cells of the first row horizontally.
        Table table = doc.FirstSection.Body.Tables[0];
        Row firstRow = table.Rows[0];
        Cell firstCell = firstRow.Cells[0];
        Cell secondCell = firstRow.Cells[1];

        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Optionally merge the third cell as well to span all three columns.
        // Uncomment the following lines to merge the third cell.
        // Cell thirdCell = firstRow.Cells[2];
        // firstCell.CellFormat.HorizontalMerge = CellMerge.First; // already set
        // thirdCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
