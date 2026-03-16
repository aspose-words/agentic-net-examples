using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageExtractionDemo
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the source ODT files.
            string sourceFolder = @"C:\InputOdt";

            // Root folder where extracted images will be placed.
            // Each document will have its own sub‑folder named after the original file (without extension).
            string destinationRoot = @"C:\ExtractedImages";

            // Ensure the destination root exists.
            Directory.CreateDirectory(destinationRoot);

            // Process every *.odt file in the source folder.
            foreach (string odtPath in Directory.GetFiles(sourceFolder, "*.odt"))
            {
                // Derive a clean name for the document (no path, no extension).
                string documentName = Path.GetFileNameWithoutExtension(odtPath);

                // Create a dedicated folder for this document's images.
                string documentImageFolder = Path.Combine(destinationRoot, documentName);
                Directory.CreateDirectory(documentImageFolder);

                // Load the ODT document. The constructor automatically detects the format.
                Document doc = new Document(odtPath);

                // Retrieve all Shape nodes (including images) from the document.
                NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;

                // Iterate through the shapes, saving only those that actually contain an image.
                foreach (Shape shape in shapes.OfType<Shape>())
                {
                    if (shape.HasImage)
                    {
                        // Determine the appropriate file extension for the image format.
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                        // Build a unique file name for the extracted image.
                        string imageFileName = $"Image_{imageIndex}{extension}";
                        string imageFullPath = Path.Combine(documentImageFolder, imageFileName);

                        // Save the image data to the file system.
                        shape.ImageData.Save(imageFullPath);

                        imageIndex++;
                    }
                }

                Console.WriteLine($"Extracted {imageIndex} image(s) from \"{odtPath}\" to \"{documentImageFolder}\".");
            }

            Console.WriteLine("Image extraction completed.");
        }
    }
}
