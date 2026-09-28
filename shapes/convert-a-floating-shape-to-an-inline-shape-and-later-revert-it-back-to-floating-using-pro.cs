using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape and configure it as a floating shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 100;
        shape.Top = 100;

        // Ensure the shape is floating (not inline).
        if (shape.WrapType == WrapType.Inline)
            throw new Exception("Shape should be floating after initial insertion.");

        // Convert the floating shape to an inline shape.
        shape.WrapType = WrapType.Inline;

        // Verify the conversion succeeded.
        if (shape.WrapType != WrapType.Inline)
            throw new Exception("Conversion to inline shape failed.");

        // Revert the shape back to floating.
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 150;
        shape.Top = 150;

        // Verify the shape is floating again.
        if (shape.WrapType == WrapType.Inline)
            throw new Exception("Reverting to floating shape failed.");

        // Save the document.
        string outputPath = "FloatingInlineFloating.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
