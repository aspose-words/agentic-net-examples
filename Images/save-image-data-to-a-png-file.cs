using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class SaveImageToPng
{
    static void Main()
    {
        // Path to the source document that contains images.
        string docsDir = @"C:\Docs\";
        string sourceFile = Path.Combine(docsDir, "Images.docx");

        // Load the document.
        Document doc = new Document(sourceFile);

        // Find the first shape that has an image.
        Shape imageShape = doc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .FirstOrDefault(s => s.HasImage);

        if (imageShape == null)
        {
            Console.WriteLine("No image found in the document.");
            return;
        }

        // Determine the appropriate file extension for the image type.
        // For this task we force PNG regardless of the original format.
        string outputFile = Path.Combine(docsDir, "ExtractedImage.png");

        // Save the image data to a PNG file.
        // ImageData.Save(string) automatically chooses the format based on the file extension.
        imageShape.ImageData.Save(outputFile);

        Console.WriteLine($"Image saved to: {outputFile}");
    }
}
