using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class WatermarkInTableCell
{
    static void Main()
    {
        // Load an existing DOT template (or create a new document if you prefer).
        Document doc = new Document("Template.dot");

        // Build a simple 1‑row, 1‑column table if the document does not already contain one.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();
        builder.InsertCell();               // Create the target cell.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first cell of the first table.
        Cell targetCell = doc.FirstSection.Body.Tables[0].Rows[0].Cells[0];

        // Move the builder's cursor to the beginning of the cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert a shape that will act as a watermark inside the cell.
        // The shape contains plain text; we set its wrapping to None and place it behind the text.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText);
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        // Font size for a TextPath is controlled by the shape's height; no FontSize property exists.
        watermarkShape.Width = 200;
        watermarkShape.Height = 50;
        watermarkShape.WrapType = WrapType.None;   // No text wrapping – behaves like a watermark.
        watermarkShape.BehindText = true;          // Render behind the cell's regular content.
        // Position the shape relative to the column (the cell) – the default vertical position works for a cell.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.Left = 0;
        watermarkShape.Top = 0;

        // Insert the shape into the cell.
        builder.InsertNode(watermarkShape);

        // Save the modified DOT document.
        doc.Save("WatermarkedTableCell.dot");
    }
}
