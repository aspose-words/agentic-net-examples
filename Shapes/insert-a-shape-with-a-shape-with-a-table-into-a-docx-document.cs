using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertShapeWithTable
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating textbox shape. The shape will act as a container for the table.
        Shape shape = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Position the shape in the middle of the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.HorizontalAlignment = HorizontalAlignment.Center;
        shape.VerticalAlignment = VerticalAlignment.Center;
        shape.WrapType = WrapType.None; // No text wrapping around the shape.

        // Move the builder's cursor inside the shape so that subsequent inserts go into it.
        builder.MoveTo(shape.FirstParagraph);

        // Build a simple 2x2 table inside the shape.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header 1");
        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell A1");
        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optionally, adjust table formatting (auto‑fit to contents).
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to a DOCX file.
        doc.Save("ShapeWithTable.docx");
    }
}
