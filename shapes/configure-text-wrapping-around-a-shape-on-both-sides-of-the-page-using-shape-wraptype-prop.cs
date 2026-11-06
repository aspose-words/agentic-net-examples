using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some introductory text.
        builder.Writeln("This paragraph appears before the shape. The text will wrap around the shape on both sides of the page.");

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        // Position the shape relative to the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = ConvertUtil.MillimeterToPoint(20);
        shape.Top = ConvertUtil.MillimeterToPoint(20);
        // Configure text wrapping: square wrapping on both sides.
        shape.WrapType = WrapType.Square;
        shape.WrapSide = WrapSide.Both;

        // Continue with more text to demonstrate wrapping.
        builder.Writeln("This is a long paragraph that follows the shape. It should flow on both the left and right sides of the rectangle, demonstrating the BothSides wrapping behavior. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, indicate success (no console input required).
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }
}
