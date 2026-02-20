using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromWordMl
{
    static void Main()
    {
        // Path to the WORDML (XML) document.
        string inputPath = @"C:\Docs\input.xml";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Load the WORDML document.
        Document doc = new Document(inputPath);

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine a suitable file extension based on the image type.
                string extension = shape.ImageData.ImageType switch
                {
                    ImageType.Jpeg => ".jpg",
                    ImageType.Png => ".png",
                    ImageType.Bmp => ".bmp",
                    ImageType.Gif => ".gif",
                    ImageType.WebP => ".webp",
                    ImageType.Emf => ".emf",
                    ImageType.Wmf => ".wmf",
                    ImageType.Pict => ".pict",
                    ImageType.Eps => ".eps",
                    _ => ".bin"
                };

                // Build a unique file name for each extracted image.
                string fileName = $"Image_{imageIndex}{extension}";
                string filePath = Path.Combine(outputFolder, fileName);

                // Write the raw image bytes to the file.
                File.WriteAllBytes(filePath, shape.ImageData.ImageBytes);

                Console.WriteLine($"Extracted image saved to: {filePath}");
                imageIndex++;
            }
        }

        Console.WriteLine("Image extraction completed.");
    }
}
