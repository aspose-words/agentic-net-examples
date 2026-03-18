using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Build a table with a merged cell ----------
        Table table = builder.StartTable();

        // First row – first cell will be the start of a horizontally and vertically merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // The next two cells are merged with the previous one (horizontal merge).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Second row – normal (unmerged) cells.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.InsertCell();
        builder.Write("Cell 2,3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // ---------- Insert a watermark shape into the merged cell ----------
        // Move the cursor to the first paragraph of the merged cell.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating text shape that will serve as the watermark.
        Shape watermark = builder.InsertShape(ShapeType.TextPlainText, 200, 50);
        watermark.WrapType = WrapType.None;                     // No text wrapping – shape stays on top.
        watermark.BehindText = true;                           // Render behind the cell contents.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        watermark.Left = 0;                                    // Align to the left edge of the cell.
        watermark.Top = 0;                                     // Align to the top edge of the cell.
        watermark.Rotation = -45;                              // Diagonal layout.

        // Configure the visual appearance of the watermark text.
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        // Removed TextPath.FontSize – not available in older Aspose.Words versions.
        watermark.TextPath.Bold = true;
        watermark.Fill.Color = Color.LightGray;                // Light gray fill makes it semi‑transparent.
        watermark.StrokeColor = Color.LightGray;

        // Save the resulting document.
        doc.Save("WatermarkInMergedCell.docx");
    }
}
