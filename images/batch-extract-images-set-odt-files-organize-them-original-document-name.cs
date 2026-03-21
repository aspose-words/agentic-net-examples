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
            // Folder that contains the source ODT files (relative to the executable's directory).
            string sourceFolder = Path.Combine(AppContext.BaseDirectory, "InputOdt");

            // Root folder where extracted images will be placed.
            // Each document will have its own sub‑folder named after the original file (without extension).
            string destinationRoot = Path.Combine(AppContext.BaseDirectory, "ExtractedImages");

            // Ensure both source and destination folders exist.
            Directory.CreateDirectory(sourceFolder);
            Directory.CreateDirectory(destinationRoot);

            // Get all *.odt files in the source folder.
            string[] odtFiles = Directory.GetFiles(sourceFolder, "*.odt");

            if (odtFiles.Length == 0)
            {
                Console.WriteLine($"No ODT files found in \"{sourceFolder}\". Place ODT files there and rerun the program.");
                return;
            }

            // Process each ODT file.
            foreach (string odtPath in odtFiles)
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
