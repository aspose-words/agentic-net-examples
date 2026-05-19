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

        // Start building a table.
        builder.StartTable();

        // ----- First row -----
        // Cell (row 1, column 1) – first cell in a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cell");

        // Cell (row 1, column 2) – regular, not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Regular cell");
        builder.EndRow();

        // ----- Second row -----
        // Cell (row 2, column 1) – merges with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text is written to a merged‑down cell.

        // Cell (row 2, column 2) – regular, not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Second row cell");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerticalMergeTable.docx");
        doc.Save(outputPath);
    }
}
