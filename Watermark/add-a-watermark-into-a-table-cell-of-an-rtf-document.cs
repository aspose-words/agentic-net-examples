using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Select the first cell where the watermark will be placed.
        Cell targetCell = table.Rows[0].Cells[0];
        builder.MoveTo(targetCell.FirstParagraph);

        // Create a floating shape that will serve as a watermark inside the cell.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText);
        watermarkShape.Width = 200;
        watermarkShape.Height = 50;
        watermarkShape.WrapType = WrapType.None;               // No text wrapping.
        watermarkShape.BehindText = true;                     // Appear behind cell contents.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        // RelativeVerticalPosition.Row is not available in the current Aspose.Words version; the default (Paragraph) works for a cell.
        watermarkShape.Left = 0;
        watermarkShape.Top = 0;

        // Configure the watermark text.
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        // FontSize property is not present in this version; size can be controlled via the shape dimensions.
        watermarkShape.TextPath.Bold = true;
        watermarkShape.FillColor = Color.LightGray;
        watermarkShape.StrokeColor = Color.LightGray;

        // Insert the shape into the document.
        builder.InsertNode(watermarkShape);

        // Save the document as RTF.
        RtfSaveOptions saveOptions = new RtfSaveOptions();
        doc.Save("TableCellWatermark.rtf", saveOptions);
    }
}
