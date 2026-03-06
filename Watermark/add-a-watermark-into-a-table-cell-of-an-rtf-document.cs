using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing RTF document.
        Document doc = new Document("Input.rtf");

        // Create a DocumentBuilder for editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the document contains at least one table.
        if (doc.GetChildNodes(NodeType.Table, true).Count == 0)
        {
            // Build a simple 1x1 table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell with watermark");
            builder.EndRow();
            builder.EndTable();
        }

        // Move the cursor to the first cell of the first table.
        builder.MoveToCell(0, 0, 0, 0);

        // Create a text shape that will act as a watermark inside the cell.
        Shape watermark = new Shape(doc, ShapeType.TextPlainText);
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        // FontSize is not available in older versions of Aspose.Words; the size of the text can be controlled by the shape's height.
        watermark.TextPath.Bold = true;

        // Make the shape transparent except for the text.
        watermark.Fill.Color = Color.Transparent;
        watermark.Stroke.Color = Color.Transparent;

        // Rotate the shape to give a typical watermark appearance.
        watermark.Rotation = -45;

        // Position the shape relative to the cell (horizontal only; vertical defaults to the cell).
        watermark.WrapType = WrapType.None;
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        // RelativeVerticalPosition.Row is not present in this version; the default positioning works for a cell.

        // Set size of the shape (height influences the text size).
        watermark.Width = 300;
        watermark.Height = 50;

        // Center the shape inside the cell.
        watermark.HorizontalAlignment = HorizontalAlignment.Center;
        watermark.VerticalAlignment = VerticalAlignment.Center;

        // Insert the watermark shape into the cell.
        builder.InsertNode(watermark);

        // Save the modified document back to RTF format.
        doc.Save("Output.rtf");
    }
}
