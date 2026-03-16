using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageExtractionBatch
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the source DOC/DOCX files.
            string inputFolder = @"C:\InputDocs";

            // Root folder where extracted images will be saved.
            string outputRoot = @"C:\ExtractedImages";

            // Ensure the output root folder exists.
            Directory.CreateDirectory(outputRoot);

            // Process each Word document in the input folder.
            foreach (string docPath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                                                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                            f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    // Load the document using the Aspose.Words Document constructor.
                    Document doc = new Document(docPath);

                    // Retrieve all Shape nodes (including images) from the document.
                    NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

                    int imageIndex = 0;

                    // Iterate through each shape that actually contains an image.
                    foreach (Shape shape in shapes.OfType<Shape>())
                    {
                        if (!shape.HasImage)
                            continue;

                        // Determine the file extension for the image (e.g., ".png").
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                        // Create a subfolder named after the image format (without the leading dot).
                        string formatFolder = Path.Combine(outputRoot, extension.TrimStart('.').ToUpperInvariant());
                        Directory.CreateDirectory(formatFolder);

                        // Build a unique file name for the extracted image.
                        string baseDocName = Path.GetFileNameWithoutExtension(docPath);
                        string imageFileName = $"{baseDocName}_Image{imageIndex}{extension}";

                        // Save the image to the appropriate subfolder.
                        string outputPath = Path.Combine(formatFolder, imageFileName);
                        shape.ImageData.Save(outputPath);

                        imageIndex++;
                    }
                }
                catch (Exception ex)
                {
                    // Log or handle errors as needed; for this example we simply write to console.
                    Console.WriteLine($"Error processing '{docPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Image extraction completed.");
        }
    }
}
