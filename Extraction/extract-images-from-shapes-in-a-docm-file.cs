using Aspose.Words;
using Aspose.Words.Drawing;
using System;
using System.IO;
using System.Linq;

class ExtractImagesFromDocm
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("Input.docm");

        // Directory where extracted images will be saved.
        string outputFolder = "ExtractedImages";
        Directory.CreateDirectory(outputFolder);

        int imageCounter = 0;

        // Iterate through all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // Process only shapes that contain an image.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = GetExtension(shape.ImageData.ImageType);

                // Build a unique file name for the extracted image.
                string filePath = Path.Combine(outputFolder, $"Image_{imageCounter}{extension}");

                // Write the image bytes to the file.
                File.WriteAllBytes(filePath, shape.ImageData.ImageBytes);

                imageCounter++;
            }
        }
    }

    // Maps Aspose.Words.ImageType values to common file extensions.
    static string GetExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg: return ".jpg";
            case ImageType.Png:  return ".png";
            case ImageType.Bmp:  return ".bmp";
            case ImageType.Gif:  return ".gif";
            case ImageType.Emf:  return ".emf";
            case ImageType.Wmf:  return ".wmf";
            case ImageType.WebP: return ".webp";
            case ImageType.Eps:  return ".eps";
            default:             return ".bin";
        }
    }
}
