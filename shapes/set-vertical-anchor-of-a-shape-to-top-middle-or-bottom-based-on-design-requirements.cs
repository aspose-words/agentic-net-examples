using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetVerticalAnchorExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first shape and set its vertical alignment to Top.
        Shape topShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        topShape.WrapType = WrapType.None; // Make the shape floating.
        topShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        topShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        topShape.VerticalAlignment = VerticalAlignment.Top; // Anchor at the top of the page.

        // Insert second shape and set its vertical alignment to Center (middle).
        Shape middleShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        middleShape.WrapType = WrapType.None;
        middleShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        middleShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        middleShape.VerticalAlignment = VerticalAlignment.Center; // Anchor at the middle of the page.

        // Insert third shape and set its vertical alignment to Bottom.
        Shape bottomShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        bottomShape.WrapType = WrapType.None;
        bottomShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        bottomShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        bottomShape.VerticalAlignment = VerticalAlignment.Bottom; // Anchor at the bottom of the page.

        // Save the document.
        string outputPath = "VerticalAnchorShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
