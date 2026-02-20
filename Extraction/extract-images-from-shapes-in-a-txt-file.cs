using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading; // Added namespace for TxtLoadOptions

class ExtractImagesFromTxt
{
    static void Main()
    {
        // Path to the source TXT file.
        string txtPath = @"C:\Docs\source.txt";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the TXT file into a Document object using TxtLoadOptions.
        Document doc = new Document(txtPath, new TxtLoadOptions());

        int imageIndex = 0;

        // Iterate through all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine a suitable file extension based on the image type.
                string extension = shape.ImageData.ImageType switch
                {
                    ImageType.Jpeg => "jpg",
                    ImageType.Png  => "png",
                    ImageType.Bmp  => "bmp",
                    ImageType.Gif  => "gif",
                    ImageType.WebP => "webp",
                    ImageType.Emf  => "emf",
                    ImageType.Wmf  => "wmf",
                    ImageType.Pict => "pict",
                    ImageType.Eps  => "eps",
                    _              => "bin"
                };

                // Build the output file name.
                string fileName = Path.Combine(outputFolder, $"Image_{imageIndex}.{extension}");

                // Save the image data to the file.
                shape.ImageData.Save(fileName);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
