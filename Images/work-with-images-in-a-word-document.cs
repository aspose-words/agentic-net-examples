using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ImageDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert an inline image from a file path.
        // -----------------------------------------------------------------
        // The image will be inserted at the current cursor position with its original size.
        Shape inlineImage = builder.InsertImage(@"C:\Images\Logo.jpg");
        // Add a paragraph break after the image.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 2. Insert a floating image with custom size and positioning.
        // -----------------------------------------------------------------
        // Parameters: file name, horizontal position relative to page margin,
        // left offset, vertical position relative to page margin, top offset,
        // width, height, wrap type.
        Shape floatingImage = builder.InsertImage(
            @"C:\Images\Watermark.png",
            RelativeHorizontalPosition.Margin, 100,
            RelativeVerticalPosition.Margin, 100,
            200, 100,
            WrapType.Square);

        // Configure the floating image to appear behind text.
        floatingImage.BehindText = true;
        // Center the image on the page.
        floatingImage.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingImage.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingImage.HorizontalAlignment = HorizontalAlignment.Center;
        floatingImage.VerticalAlignment = VerticalAlignment.Center;

        // -----------------------------------------------------------------
        // 3. Create a shape and set its image using ImageData.SetImage.
        // -----------------------------------------------------------------
        Shape shapeWithImage = new Shape(doc, ShapeType.Image);
        shapeWithImage.Width = 150;
        shapeWithImage.Height = 150;
        // Load an image from a stream and assign it to the shape.
        using (FileStream imgStream = new FileStream(@"C:\Images\Sample.png", FileMode.Open, FileAccess.Read))
        {
            shapeWithImage.ImageData.SetImage(imgStream);
        }
        // Insert the shape into the document.
        builder.InsertNode(shapeWithImage);
        builder.Writeln();

        // -----------------------------------------------------------------
        // 4. Save the document to a DOCX file.
        // -----------------------------------------------------------------
        doc.Save(@"C:\Output\ImageDemo.docx", SaveFormat.Docx);
    }
}
