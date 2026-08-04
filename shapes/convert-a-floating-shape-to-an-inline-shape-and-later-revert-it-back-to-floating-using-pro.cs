using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeConversionExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape floatingShape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left position
            RelativeVerticalPosition.Page, 100,     // top position
            100, 100,                               // width, height
            WrapType.None);                         // floating (no wrap)

        // Validate that the shape is floating.
        if (floatingShape.IsInline)
            throw new InvalidOperationException("Shape should be floating after insertion.");

        // Convert the floating shape to an inline shape.
        floatingShape.WrapType = WrapType.Inline;

        // Validate that the shape is now inline.
        if (!floatingShape.IsInline)
            throw new InvalidOperationException("Shape should be inline after conversion.");

        // Revert the shape back to floating.
        floatingShape.WrapType = WrapType.None;
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.Left = 100;
        floatingShape.Top = 100;
        floatingShape.Width = 100;
        floatingShape.Height = 100;

        // Validate that the shape is floating again.
        if (floatingShape.IsInline)
            throw new InvalidOperationException("Shape should be floating after reverting.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeConversion.docx");
        doc.Save(outputPath);

        // Simple confirmation (no interactive prompts).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
