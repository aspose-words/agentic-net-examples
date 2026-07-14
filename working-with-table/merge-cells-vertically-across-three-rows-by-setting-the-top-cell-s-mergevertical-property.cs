using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. The builder will automatically create the first row when we insert the first cell.
        Table table = builder.StartTable();

        // ---------- Row 1 ----------
        // First cell (column 0) – this cell will be the start of a vertical merge spanning three rows.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First; // Mark as the first cell in a vertical merge.
        builder.Write("Merged vertically across three rows");

        // Second cell (column 1) – regular content.
        builder.InsertCell();
        builder.Write("Row 1, Column 2");
        builder.EndRow();

        // ---------- Row 2 ----------
        // First cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue the vertical merge.
        // No text is written for merged cells; they must remain empty.

        // Second cell – regular content.
        builder.InsertCell();
        builder.Write("Row 2, Column 2");
        builder.EndRow();

        // ---------- Row 3 ----------
        // First cell – continue the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell – regular content.
        builder.InsertCell();
        builder.Write("Row 3, Column 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(outputPath);
    }
}
