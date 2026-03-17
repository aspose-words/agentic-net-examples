using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will be the anchor for the shape.
        builder.Writeln("Paragraph before the shape.");

        // Insert a rectangle shape. The InsertShape overload that takes only width/height
        // is used, then the positioning properties are set manually.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Position the shape relative to the preceding paragraph.
        // Horizontal positioning relative to the column (the closest analogue to "paragraph" in older versions).
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        shape.Left = 20; // 20 points from the left edge of the column/paragraph.
        // Vertical positioning relative to the paragraph.
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        shape.Top = 10; // 10 points below the paragraph anchor.

        // Set wrapping and visual formatting.
        shape.WrapType = WrapType.Square;
        shape.FillColor = Color.LightBlue;
        shape.Stroke.Color = Color.DarkBlue;

        // Continue adding content after the shape.
        builder.Writeln("Paragraph after the shape.");

        // Save the document.
        doc.Save("ShapeRelativeToParagraph.docx");
    }
}
