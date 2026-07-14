using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class FloatingShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        // Parameters: shape type, horizontal position reference, left, vertical position reference, top,
        // width, height, wrap type (None = floating).
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            100, 50,
            WrapType.None);

        // Verify the shape is floating.
        if (shape.IsInline)
            throw new Exception("Shape should be floating after insertion.");

        // Convert the floating shape to an inline shape.
        shape.WrapType = WrapType.Inline;

        // Verify the shape is now inline.
        if (!shape.IsInline)
            throw new Exception("Shape should be inline after conversion.");

        // Convert the inline shape back to a floating shape.
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 100;   // Position from the left edge of the page.
        shape.Top = 100;    // Position from the top edge of the page.
        shape.Width = 100;
        shape.Height = 50;

        // Verify the shape is floating again.
        if (shape.IsInline)
            throw new Exception("Shape should be floating after reverting.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingInlineFloating.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the output document.");

        // Inform that the process completed successfully.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
