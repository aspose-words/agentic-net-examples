using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using SkiaSharp;

class Program
{
    static void Main()
    {
        // Directories for input images and output files.
        string dataDir = @"C:\Images\";
        string outputDir = @"C:\Output\";
        Directory.CreateDirectory(outputDir);

        // 1. Create a new document and insert an inline image from a file.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(Path.Combine(dataDir, "Logo.jpg"));
        builder.Writeln();

        // 2. Insert a floating image with custom size, position and wrap type.
        Shape floating = builder.InsertImage(
            Path.Combine(dataDir, "Logo.jpg"),
            RelativeHorizontalPosition.Margin, 100,          // left offset
            RelativeVerticalPosition.Margin, 100,            // top offset
            200, 100,                                        // width, height (points)
            WrapType.Square);                                // text wrap
        floating.BehindText = false; // keep the image in front of text.

        builder.Writeln();

        // 3. Insert a rectangle shape and fill it with an image.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rect.Fill.SetImage(Path.Combine(dataDir, "Logo.jpg"));
        rect.WrapType = WrapType.None;
        rect.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        rect.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        rect.HorizontalAlignment = HorizontalAlignment.Center;
        rect.VerticalAlignment = VerticalAlignment.Center;

        builder.Writeln();

        // 4. Add an image watermark to the whole document.
        // Aspose.Words for .NET uses SkiaSharp.SKBitmap for image watermarks.
        using (SKBitmap watermarkBitmap = SKBitmap.Decode(Path.Combine(dataDir, "Transparent background logo.png")))
        {
            doc.Watermark.SetImage(watermarkBitmap);
        }

        // 5. Save the document.
        string outFile = Path.Combine(outputDir, "ImageDemo.docx");
        doc.Save(outFile);

        // 6. Load the saved document and extract all embedded images.
        Document loaded = new Document(outFile);
        NodeCollection shapes = loaded.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imagePath = Path.Combine(outputDir, $"Extracted_{imageIndex}{extension}");
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }
        }
    }
}
