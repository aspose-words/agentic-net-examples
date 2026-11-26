using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape (initially inline).
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Convert the shape to floating by setting wrapping and positioning.
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 100;
        shape.Top = 100;

        // Save the document with the floating shape.
        const string floatingPath = "FloatingShape.docx";
        doc.Save(floatingPath);
        ValidateFileExists(floatingPath);

        // Convert the floating shape back to inline.
        shape.WrapType = WrapType.Inline;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin; // reset to default positioning
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;   // reset to default positioning
        shape.Left = 0;
        shape.Top = 0;

        // Save the document with the inline shape.
        const string inlinePath = "InlineShape.docx";
        doc.Save(inlinePath);
        ValidateFileExists(inlinePath);

        // Revert the shape to floating again.
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 150;
        shape.Top = 150;

        // Save the document with the shape floating again.
        const string floatingAgainPath = "FloatingAgain.docx";
        doc.Save(floatingAgainPath);
        ValidateFileExists(floatingAgainPath);
    }

    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
            throw new InvalidOperationException($"The file '{path}' was not created.");
    }
}
