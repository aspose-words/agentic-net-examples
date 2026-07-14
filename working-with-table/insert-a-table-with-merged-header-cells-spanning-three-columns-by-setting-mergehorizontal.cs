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

        // ---- Header row (merged across three columns) ----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning three columns");

        // Insert the next two cells and merge them with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the header row.
        builder.EndRow();

        // Reset merge settings for normal cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ---- First data row (three separate cells) ----
        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.InsertCell();
        builder.Write("Row 1, Col 3");
        builder.EndRow();

        // ---- Second data row (three separate cells) ----
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.InsertCell();
        builder.Write("Row 2, Col 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedHeaderTable.docx");
        doc.Save(outputPath);
    }
}
