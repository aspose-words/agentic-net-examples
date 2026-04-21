using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Add a text watermark to the document (this creates a shape in the header).
        doc.Watermark.SetText("Cell Watermark");

        // Locate the watermark shape that was just added.
        // Watermarks are placed in the primary header of the first section.
        HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Shape watermarkShape = (Shape)header.GetChild(NodeType.Shape, 0, true);

        // Remove the shape from the header using the shape's own Remove method.
        watermarkShape.Remove();

        // Insert the watermark shape into a specific table cell (e.g., second cell of first row).
        Table table = doc.FirstSection.Body.Tables[0];
        Cell targetCell = table.Rows[0].Cells[1];

        // Adjust positioning so the shape is no longer anchored to the header.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        watermarkShape.WrapType = WrapType.None;
        watermarkShape.BehindText = true;

        // Append the shape to the cell's first paragraph.
        Paragraph para = targetCell.FirstParagraph;
        para.AppendChild(watermarkShape);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellWatermark.docx");
        doc.Save(outputPath);
    }
}
