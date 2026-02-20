using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromShapes
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Folder where extracted images will be saved.
        string outputFolder = "ExtractedImages";
        Directory.CreateDirectory(outputFolder);

        // Load the document.
        Document doc = new Document(inputPath);

        int imageIndex = 0;

        // Iterate through all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that contain an image.
            if (shape.HasImage)
            {
                // Determine the image format to create an appropriate file extension.
                string extension = GetExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string outputPath = Path.Combine(outputFolder, $"Image_{imageIndex}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(outputPath);

                Console.WriteLine($"Saved image {imageIndex} to {outputPath}");
                imageIndex++;
            }
        }

        Console.WriteLine("Image extraction completed.");
    }

    // Helper method to map Aspose.Words.ImageType to a file extension.
    private static string GetExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpeg";
            case ImageType.Png:
                return ".png";
            case ImageType.Bmp:
                return ".bmp";
            case ImageType.Gif:
                return ".gif";
            case ImageType.Emf:
                return ".emf";
            case ImageType.Wmf:
                return ".wmf";
            case ImageType.Pict:
                return ".pict";
            case ImageType.Eps:
                return ".eps";
            case ImageType.WebP:
                return ".webp";
            default:
                return ".bin";
        }
    }
}
