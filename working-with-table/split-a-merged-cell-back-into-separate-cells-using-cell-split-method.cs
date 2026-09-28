using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitMergedCellExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with two cells in a single row.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Merged Cell");
        builder.InsertCell(); // Placeholder for the second cell.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];
        Row row = table.Rows[0];
        Cell firstCell = row.Cells[0];
        Cell secondCell = row.Cells[1];

        // Merge the two cells horizontally.
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ----- Split the merged cells back into separate cells -----
        // Aspose.Words does not provide a Cell.Split method. To "split" we simply
        // clear the merge settings on both cells, which restores them as independent cells.
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;
        // -----------------------------------------------------------

        // Write distinct text into the now separate cells.
        builder.MoveTo(firstCell.FirstParagraph);
        builder.Writeln("Cell 1");
        builder.MoveTo(secondCell.FirstParagraph);
        builder.Writeln("Cell 2");

        // Save the document.
        string outputPath = "SplitMergedCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
