// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkInTableCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a table with a single cell to the document.
        Table table = new Table(doc);
        Row row = new Row(doc);
        Cell cell = new Cell(doc);
        row.Cells.Add(cell);
        table.Rows.Add(row);
        doc.FirstSection.Body.AppendChild(table);

        // Create a shape that will act as a watermark inside the cell.
        // The shape is a text box with transparent background.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText);
        watermarkShape.Width = 300;   // Adjust size as needed.
        watermarkShape.Height = 100;
        watermarkShape.WrapType = WrapType.Inline; // Keep it inside the cell flow.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Set the text of the watermark.
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        watermarkShape.TextPath.FontSize = 36;
        watermarkShape.TextPath.Bold = true;
        watermarkShape.TextPath.FillColor = Color.Red;

        // Rotate the shape to give a typical watermark appearance.
        watermarkShape.Rotation = -45;

        // Make the shape semi‑transparent.
        watermarkShape.Fill.Color = Color.FromArgb(50, Color.White);
        watermarkShape.StrokeColor = Color.Transparent;

        // Insert the shape into the cell.
        cell.FirstParagraph.AppendChild(watermarkShape);

        // Save the document as PDF.
        doc.Save("WatermarkInCell.pdf");
    }
}
