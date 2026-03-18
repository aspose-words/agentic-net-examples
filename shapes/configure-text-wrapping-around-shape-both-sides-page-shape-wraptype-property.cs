using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 150);

        // Make the shape floating and wrap text around its bounding box.
        shape.WrapType = WrapType.Square;

        // Configure the text to wrap on both sides of the shape.
        shape.WrapSide = WrapSide.Both;

        // Position the shape relative to the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 100;   // distance from the left edge of the page
        shape.Top = 100;    // distance from the top edge of the page

        // Add some surrounding text to demonstrate wrapping.
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi " +
            "ut aliquip ex ea commodo consequat.");

        // Save the document to a file.
        doc.Save("ShapeWrapBothSides.docx");
    }
}
