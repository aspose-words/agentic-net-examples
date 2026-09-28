using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the shape.
        builder.Writeln("This is some text before the shape. It will wrap around the shape on both sides of the page.");

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
        // Configure wrapping: square wrapping on both sides.
        shape.WrapType = WrapType.Square;
        shape.WrapSide = WrapSide.Both; // Correct enum value
        // Position the shape relative to the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = ConvertUtil.MillimeterToPoint(20);
        shape.Top = ConvertUtil.MillimeterToPoint(20);
        // Optional visual styling.
        shape.StrokeColor = Color.Black;
        shape.FillColor = Color.LightBlue;

        // Add more text after the shape.
        builder.Writeln();
        builder.Writeln("This is some text after the shape. It should also wrap around the shape on both sides, demonstrating the Both setting.");

        // Save the document.
        string outputPath = "WrapShapeBothSides.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
