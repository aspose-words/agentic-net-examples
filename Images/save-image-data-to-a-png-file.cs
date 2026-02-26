using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load a document that contains images.
        Document doc = new Document("Images.docx");

        // Retrieve the first shape that actually holds an image.
        Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (imageShape == null || !imageShape.HasImage)
        {
            Console.WriteLine("No image found in the document.");
            return;
        }

        // Save the image data to a PNG file using the ImageData.Save(string) method.
        string outputFile = "ExtractedImage.png";
        imageShape.ImageData.Save(outputFile);

        Console.WriteLine($"Image successfully saved to: {outputFile}");
    }
}
