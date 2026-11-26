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
        // Row 1: first cell merges horizontally across two columns.
        builder.StartTable();

        // Cell (0,0) – start of horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Horizontal Merge");

        // Cell (0,1) – continuation of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged part.
        builder.Writeln(string.Empty);

        // End first row.
        builder.EndRow();

        // Reset merge settings for subsequent cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // Row 2: first cell merges vertically across two rows.
        // Cell (1,0) – start of vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Writeln("Vertical Merge");

        // Cell (1,1) – normal cell.
        builder.InsertCell();
        builder.Writeln("Cell B2");

        // End second row.
        builder.EndRow();

        // Reset vertical merge for the next row.
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // Row 3: continuation of the vertical merge from above.
        // Cell (2,0) – continuation of vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged part.
        builder.Writeln(string.Empty);

        // Cell (2,1) – normal cell.
        builder.InsertCell();
        builder.Writeln("Cell C3");

        // End third row and the table.
        builder.EndRow();
        builder.EndTable();

        // Prepare HTML save options to produce plain HTML with proper colspan/rowspan.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Inline CSS keeps the output self‑contained while still being plain HTML.
            CssStyleSheetType = CssStyleSheetType.Inline,
            // Produce readable formatting.
            PrettyFormat = true
        };

        // Define output path.
        string outputPath = "output.html";

        // Save the document as HTML.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the HTML output file.");

        Console.WriteLine($"Table successfully converted to HTML. File saved at: {Path.GetFullPath(outputPath)}");
    }
}
