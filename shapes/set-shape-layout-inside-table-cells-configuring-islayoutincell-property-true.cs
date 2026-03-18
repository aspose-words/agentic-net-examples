using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ShapeLayoutInTableCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // Move the builder's cursor to the first paragraph of the current cell.
        // This ensures that the shape will be placed inside this cell.
        builder.MoveTo(builder.CurrentParagraph);

        // Insert a floating rectangle shape.
        // Width = 100 points, Height = 100 points, no text wrapping (WrapType.None).
        // The overload that includes positioning parameters must be used for floating shapes.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Margin, 0,   // left position relative to the left margin
            RelativeVerticalPosition.Margin, 0,     // top position relative to the top margin
            100, 100,                               // width, height
            WrapType.None);

        // Configure the shape to be displayed inside the table cell.
        // The shape must be a floating shape (WrapType != Inline) for this property to take effect.
        shape.IsLayoutInCell = true;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        doc.Save("Shape.LayoutInTableCell.docx");
    }
}
