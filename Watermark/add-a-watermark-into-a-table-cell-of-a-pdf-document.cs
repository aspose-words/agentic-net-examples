using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class TableCellWatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Build a simple 1x1 table and obtain a reference to its single cell.
        // -----------------------------------------------------------------
        Table table = builder.StartTable();
        builder.InsertCell();                     // create the first (and only) cell
        builder.Write("Cell content");            // any regular content
        // The builder's current paragraph resides inside the cell we just created.
        Cell targetCell = (Cell)builder.CurrentParagraph.ParentNode;

        // -----------------------------------------------------------------
        // Insert a text shape that will act as a watermark inside the cell.
        // -----------------------------------------------------------------
        // Move the cursor to the beginning of the cell so the shape is placed there.
        builder.MoveTo(targetCell.FirstParagraph);
        // Create a plain‑text shape (WordArt) that will contain the watermark text.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText);
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        // FontSize property is not available in older Aspose.Words versions; size can be controlled via Width/Height.
        watermarkShape.Width = 300;               // adjust size as needed
        watermarkShape.Height = 100;
        // Rotate the shape to give a typical watermark appearance.
        watermarkShape.Rotation = -45f;            // degrees (use Rotation instead of RotationAngle)
        // Place the shape behind the cell text and prevent text wrapping.
        watermarkShape.WrapType = WrapType.None;
        watermarkShape.BehindText = true;
        // Insert the shape into the document at the current cursor position.
        builder.InsertNode(watermarkShape);

        // Finish the table.
        builder.EndRow();
        builder.EndTable();

        // -----------------------------------------------------------------
        // Save the document as a PDF file.
        // -----------------------------------------------------------------
        doc.Save("TableCellWatermark.pdf", SaveFormat.Pdf);
    }
}
