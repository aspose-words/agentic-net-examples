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

        // ---------- First row ----------
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // ---------- Second row ----------
        // Insert the first cell of the second row and mark it as the start of a horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell (Row 2, Cells 1+2)");

        // Insert the second cell of the second row and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text is needed for the merged cell.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "MergedTable.docx");
        doc.Save(outputPath);
    }
}
