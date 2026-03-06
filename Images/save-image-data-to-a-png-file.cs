using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class SaveImageToPng
{
    static void Main()
    {
        // Path to the source Word document that contains images.
        string docPath = @"C:\Input\Images.docx";

        // Load the document.
        Document doc = new Document(docPath);

        // Find the first shape that has an image.
        Shape imageShape = doc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .FirstOrDefault(s => s.HasImage);

        if (imageShape == null)
        {
            Console.WriteLine("No image found in the document.");
            return;
        }

        // Get the ImageData object from the shape.
        ImageData imgData = imageShape.ImageData;

        // Determine a file name with PNG extension.
        // The ImageType property tells us the original format; we force PNG.
        string outputPath = @"C:\Output\ExtractedImage.png";

        // Save the image data directly to a PNG file.
        // The ImageData.Save(string) overload automatically chooses the format
        // based on the file extension.
        imgData.Save(outputPath);

        Console.WriteLine($"Image saved to: {outputPath}");
    }
}
