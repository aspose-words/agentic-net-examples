using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file that will be inserted.
        string imagePath = "ImageDir/Logo.jpg";

        // Insert the image inline at the current cursor position.
        // This uses the DocumentBuilder.InsertImage(string) overload.
        Shape insertedImage = builder.InsertImage(imagePath);

        // The returned Shape can be further customized if needed.
        // For example, you could change its width/height:
        // insertedImage.Width = ConvertUtil.PixelToPoint(200);
        // insertedImage.Height = ConvertUtil.PixelToPoint(120);

        // Save the document to disk.
        doc.Save("Output/InlineImage.docx");
    }
}
