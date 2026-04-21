using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeConversionExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape floatingShape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left distance from page
            RelativeVerticalPosition.Page, 100,     // top distance from page
            150, 100,                               // width, height
            WrapType.None);                         // floating (no wrap)

        // Verify that the shape is initially floating.
        if (floatingShape.IsInline)
            throw new InvalidOperationException("Shape should be floating after insertion.");

        // Save the document with the floating shape.
        string floatingPath = Path.Combine(outputDir, "FloatingShape.docx");
        doc.Save(floatingPath);

        // Convert the floating shape to an inline shape.
        floatingShape.WrapType = WrapType.Inline;

        // Verify conversion to inline.
        if (!floatingShape.IsInline)
            throw new InvalidOperationException("Shape conversion to inline failed.");

        // Save the document with the inline shape.
        string inlinePath = Path.Combine(outputDir, "InlineShape.docx");
        doc.Save(inlinePath);

        // Revert the shape back to floating.
        floatingShape.WrapType = WrapType.None;
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.Left = 200;   // new left position
        floatingShape.Top = 150;    // new top position

        // Verify that the shape is floating again.
        if (floatingShape.IsInline)
            throw new InvalidOperationException("Shape conversion back to floating failed.");

        // Save the final document.
        string revertedPath = Path.Combine(outputDir, "RevertedFloatingShape.docx");
        doc.Save(revertedPath);
    }
}
