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
        // Parameters: shape type, horizontal position reference, left, vertical position reference, top, width, height, wrap type.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            100, 100,
            WrapType.None); // Floating shape (WrapType.None)

        // Verify that the shape is floating.
        if (shape.IsInline)
            throw new InvalidOperationException("Shape should be floating after insertion.");

        // Convert the floating shape to an inline shape.
        shape.WrapType = WrapType.Inline;               // Inline shapes use WrapType.Inline
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column; // Inline positioning defaults
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        shape.Left = 0;                                 // Position is ignored for inline shapes
        shape.Top = 0;

        // Verify conversion to inline.
        if (!shape.IsInline)
            throw new InvalidOperationException("Shape should be inline after conversion.");

        // Convert the inline shape back to a floating shape.
        shape.WrapType = WrapType.None;                 // Set to floating
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 150;                               // New floating position
        shape.Top = 150;

        // Verify conversion back to floating.
        if (shape.IsInline)
            throw new InvalidOperationException("Shape should be floating after reverting.");

        // Save the document to verify the result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeConversion.docx");
        doc.Save(outputPath);
    }
}
