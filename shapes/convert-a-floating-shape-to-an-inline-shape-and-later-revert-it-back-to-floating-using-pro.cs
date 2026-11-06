using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeConversionExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape. By default InsertShape creates an inline shape,
        // so we immediately change its properties to make it floating.
        Shape floatingShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        floatingShape.WrapType = WrapType.Square;                         // Floating wrap.
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.Left = 100; // Position from the left edge of the page.
        floatingShape.Top = 100;  // Position from the top edge of the page.

        // Save the document that contains the floating shape.
        string floatingPath = "FloatingShape.docx";
        doc.Save(floatingPath);
        ValidateFileExists(floatingPath, "floating shape document");

        // ----- Convert the floating shape to an inline shape -----
        // Making a shape inline is done by setting its WrapType to Inline.
        floatingShape.WrapType = WrapType.Inline;

        // Save the document that now contains the shape as inline.
        string inlinePath = "InlineShape.docx";
        doc.Save(inlinePath);
        ValidateFileExists(inlinePath, "inline shape document");

        // ----- Convert the inline shape back to a floating shape -----
        floatingShape.WrapType = WrapType.Square;                         // Restore floating wrap.
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.Left = 150; // New horizontal position.
        floatingShape.Top = 150;  // New vertical position.

        // Save the document with the restored floating shape.
        string restoredPath = "RestoredFloatingShape.docx";
        doc.Save(restoredPath);
        ValidateFileExists(restoredPath, "restored floating shape document");
    }

    private static void ValidateFileExists(string path, string description)
    {
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create the {description} at '{path}'.");
    }
}
