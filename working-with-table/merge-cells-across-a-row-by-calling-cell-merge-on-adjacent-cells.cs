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

        // Start a table.
        builder.StartTable();

        // Insert the first cell of the row and write some text.
        Cell firstCell = builder.InsertCell();
        // Mark this cell as the first cell in a horizontally merged range.
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This cell will be merged with the next one.");

        // Insert the second cell of the same row.
        Cell secondCell = builder.InsertCell();
        // Mark this cell as merged to the previous cell.
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text is written to the merged cell.

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
