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

        // Start a new table.
        Table table = builder.StartTable();

        // ----- First row (merged cell) -----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged across first row");

        // Insert additional cells in the same row and merge them with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // No text needed.

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // No text needed.

        // End the first row.
        builder.EndRow();

        // Reset merge setting for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ----- Second row (regular cells) -----
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
