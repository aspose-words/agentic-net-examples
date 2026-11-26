using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 100);

        // Make the shape floating so that rotation is applied.
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Rotate the shape by 45 degrees using the correct property.
        shape.Rotation = 45;

        // Save the document.
        const string docPath = "RotatedShape.docx";
        doc.Save(docPath);

        // Render the first page to PNG to verify the visual result.
        const string pngPath = "RotatedShape.png";
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // render only the first page
        };
        doc.Save(pngPath, options);

        // Validate that the files were created.
        if (!File.Exists(docPath))
            throw new Exception("The document file was not saved.");

        if (!File.Exists(pngPath))
            throw new Exception("The rendered image file was not saved.");

        // Validate that the shape's rotation angle is set correctly.
        if (Math.Abs(shape.Rotation - 45) > 0.001)
            throw new Exception($"Shape rotation is {shape.Rotation}, expected 45.");

        // Ensure the rendered image is not empty.
        if (new FileInfo(pngPath).Length == 0)
            throw new Exception("The rendered image file is empty.");

        Console.WriteLine("Shape rotation applied and document rendered successfully.");
    }
}
