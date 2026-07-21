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

        // Build a simple 2‑row, 3‑column table.
        builder.StartTable();

        // First row – three separate cells.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row – three cells that we will merge the first two.
        builder.InsertCell();                     // Cell that will become the first in the merged range.
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell (R2C1‑C2)");

        builder.InsertCell();                     // Cell that merges with the previous one.
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged‑into cell.

        builder.InsertCell();                     // Third cell remains independent.
        builder.Write("R2C3");
        builder.EndRow();

        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");
        doc.Save(outputPath);
    }
}
