using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertCellWatermark
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple 2x2 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Move the cursor to the target cell (first section, first table, first row, second column).
        builder.MoveToCell(0, 0, 0, 1);

        // Define watermark options (optional – you can customize appearance here).
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 24,
            Color = System.Drawing.Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply a text watermark to the document.
        // Note: Aspose.Words applies watermarks at the document level; this will be visible on all pages.
        // The API does not support attaching a watermark directly to a single table cell.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the resulting document.
        doc.Save("CellWatermark.docx");
    }
}
