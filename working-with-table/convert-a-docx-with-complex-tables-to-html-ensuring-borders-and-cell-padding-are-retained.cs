using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a complex table with merged cells, borders and padding.
        Table table = builder.StartTable();

        // First row – a header that spans two columns.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Header spanning two columns");
        // Merge the first cell with the next cell horizontally.
        Cell mergedHeader = builder.CurrentParagraph.ParentNode as Cell;
        mergedHeader.CellFormat.HorizontalMerge = CellMerge.First;
        builder.InsertCell();
        mergedHeader = builder.CurrentParagraph.ParentNode as Cell;
        mergedHeader.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.EndRow();

        // Second row – normal cells with padding.
        builder.InsertCell();
        builder.CellFormat.SetPaddings(10, 5, 10, 5); // left, top, right, bottom
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.CellFormat.SetPaddings(10, 5, 10, 5);
        builder.Write("Cell B1");
        builder.EndRow();

        // Third row – a cell that spans two rows vertically.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cell");
        builder.InsertCell();
        builder.Write("Cell B2");
        builder.EndRow();

        // Fourth row – continuation of the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No content needed for the merged cell.
        builder.InsertCell();
        builder.Write("Cell B3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply borders to the whole table.
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // Set uniform padding for the table (applies to all cells that don't have individual padding).
        table.LeftPadding = 8;
        table.RightPadding = 8;
        table.TopPadding = 4;
        table.BottomPadding = 4;

        // Prepare HTML save options to preserve borders and padding.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export all width information; this keeps the layout as close as possible.
            TableWidthOutputMode = HtmlElementSizeOutputMode.All,
            // Preserve negative indents if any (not used here but safe to enable).
            AllowNegativeIndent = true
        };

        // Define output paths.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string htmlPath = Path.Combine(outputDir, "ComplexTable.html");

        // Save the document as HTML.
        doc.Save(htmlPath, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Optionally, you could also save the original DOCX for reference.
        string docxPath = Path.Combine(outputDir, "ComplexTable.docx");
        doc.Save(docxPath);
    }
}
