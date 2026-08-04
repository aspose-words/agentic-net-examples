using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains a complex table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the outer table.
        Table outerTable = builder.StartTable();

        // -----------------------------------------------------------------
        // Row 1 – a merged cell that spans two columns.
        // -----------------------------------------------------------------
        builder.InsertCell();
        // Add some padding to this specific cell.
        builder.CellFormat.SetPaddings(5, 5, 5, 5);
        // Mark the start of a horizontal merge.
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning two columns");

        builder.InsertCell();
        // Mark the continuation of the merge.
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No content needed for the merged part.
        builder.EndRow();

        // -----------------------------------------------------------------
        // Row 2 – normal cells.
        // -----------------------------------------------------------------
        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        // -----------------------------------------------------------------
        // Row 3 – first cell contains a nested table.
        // -----------------------------------------------------------------
        builder.InsertCell();

        // Start nested table.
        Table nestedTable = builder.StartTable();

        // Nested table row 1.
        builder.InsertCell();
        builder.Write("Nested 1");
        builder.InsertCell();
        builder.Write("Nested 2");
        builder.EndRow();

        // Apply formatting to the nested table after it has at least one row.
        nestedTable.SetBorders(LineStyle.Single, 0.5, Color.Blue);
        nestedTable.LeftPadding = 2;
        nestedTable.RightPadding = 2;
        nestedTable.TopPadding = 2;
        nestedTable.BottomPadding = 2;

        // Finish nested table.
        builder.EndTable();

        // Continue outer table – second cell of the same row.
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        // Finish the outer table.
        builder.EndTable();

        // Apply formatting to the outer table after it has rows.
        outerTable.SetBorder(BorderType.Left,   LineStyle.Single, 1.5, Color.Black, true);
        outerTable.SetBorder(BorderType.Right,  LineStyle.Single, 1.5, Color.Black, true);
        outerTable.SetBorder(BorderType.Top,    LineStyle.Single, 1.5, Color.Black, true);
        outerTable.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);
        outerTable.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Gray, true);
        outerTable.SetBorder(BorderType.Vertical,   LineStyle.Single, 1.0, Color.Gray, true);

        outerTable.LeftPadding   = 10;
        outerTable.RightPadding  = 10;
        outerTable.TopPadding    = 5;
        outerTable.BottomPadding = 5;

        // Save the sample DOCX to disk.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleComplexTable.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and convert it to HTML while preserving borders
        //    and cell padding.
        // -----------------------------------------------------------------
        Document loadDoc = new Document(docPath);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export all width information (absolute and relative) to keep layout.
            TableWidthOutputMode = HtmlElementSizeOutputMode.All
        };

        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "ComplexTable.html");
        loadDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 3. Simple validation – ensure the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new Exception("HTML conversion failed: output file was not created.");
    }
}
