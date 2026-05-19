using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with complex merged cells.
        // -------------------------------------------------
        // Row 1: Cell A spans two columns (horizontal merge).
        // Row 2: Cell B spans two rows (vertical merge) and also spans two columns.
        // Row 3: Regular cells.
        // -------------------------------------------------
        Table table = builder.StartTable();

        // ----- Row 1 -----
        // First cell – start horizontal merge (colspan = 2).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning two columns");

        // Second cell – merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged cell.
        builder.EndRow();

        // Reset merge flags for subsequent cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // ----- Row 2 -----
        // First cell – start vertical merge (rowspan = 2) and also horizontal merge (colspan = 2).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Cell spanning two rows and two columns");

        // Second cell – merged horizontally with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed.
        builder.EndRow();

        // Reset merge flags again.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // ----- Row 3 -----
        // Regular cells (no merging).
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure the table's merged cells are represented by merge flags.
        // This is required for correct HTML export.
        table.ConvertToHorizontallyMergedCells();

        // Save the document as a DOCX (optional, demonstrates lifecycle compliance).
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(docxPath);

        // Export the document (containing the table) to plain HTML.
        // Aspose.Words automatically adds appropriate colspan and rowspan attributes.
        string html = doc.ToString(SaveFormat.Html);

        // Write the HTML to a file.
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.html");
        File.WriteAllText(htmlPath, html);

        // Simple validation – ensure the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML export failed.");

        // The program finishes without waiting for user input.
    }
}
