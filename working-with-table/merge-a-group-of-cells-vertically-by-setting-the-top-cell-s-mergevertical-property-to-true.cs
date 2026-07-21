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

        // Start a table with two columns.
        Table table = builder.StartTable();

        // ---------- First row ----------
        // First cell – this will be the top cell of the vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First; // Mark as the first cell in a vertical merge.
        builder.Write("Vertically merged cells");

        // Second cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // ---------- Second row ----------
        // First cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Merge with the previous (top) cell.
        // No text is written to merged cells other than the first one.

        // Second cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // The program ends here without waiting for user input.
    }
}
