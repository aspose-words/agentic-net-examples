using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell – regular content.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second cell – will contain the watermark.
        builder.InsertCell();

        // Create a TextPlainText shape that will act as a watermark (WordArt).
        Shape watermark = new Shape(doc, ShapeType.TextPlainText);
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        watermark.Width = 200;
        watermark.Height = 50;
        watermark.Rotation = -45; // Diagonal appearance.
        watermark.FillColor = Color.Transparent;
        watermark.StrokeColor = Color.Transparent;
        watermark.WrapType = WrapType.None;

        // Position the shape in the centre of the cell.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        // Use the default vertical positioning (Paragraph) – sufficient for a cell.
        watermark.HorizontalAlignment = HorizontalAlignment.Center;
        watermark.VerticalAlignment = VerticalAlignment.Center;

        // Insert the watermark shape into the current cell.
        builder.InsertNode(watermark);

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document as PDF.
        doc.Save("TableCellWatermark.pdf");
    }
}
