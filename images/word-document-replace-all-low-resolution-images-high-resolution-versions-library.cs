using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageReplacementExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source document, the folder that contains high‑resolution images,
            // and the destination document.
            string sourceDocPath = @"C:\Docs\Input.docx";
            string highResImagesFolder = @"C:\Docs\HighResImages";
            string outputDocPath = @"C:\Docs\Output.docx";

            // Load the existing Word document.
            Document doc = new Document(sourceDocPath);

            // Retrieve all Shape nodes (including pictures) from the document.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            // Define a simple threshold for what we consider a low‑resolution image.
            // Here we use the image byte size; adjust as needed for your scenario.
            const int lowResolutionByteThreshold = 50000; // 50 KB

            foreach (Shape shape in shapeNodes)
            {
                // Process only image shapes.
                if (!shape.IsImage)
                    continue;

                // If the image data is smaller than the threshold, treat it as low‑resolution.
                if (shape.ImageData.ImageBytes.Length < lowResolutionByteThreshold)
                {
                    // Use the shape's Title (or AlternativeText) as a key to locate the
                    // corresponding high‑resolution image in the library folder.
                    // This assumes that the low‑resolution image was given a meaningful title.
                    string imageKey = shape.Title;

                    if (string.IsNullOrEmpty(imageKey))
                        continue; // No key – cannot locate a replacement.

                    string highResImagePath = Path.Combine(highResImagesFolder, imageKey);

                    // Replace the image only if a matching high‑resolution file exists.
                    if (File.Exists(highResImagePath))
                    {
                        // Set the new image for the shape.
                        shape.ImageData.SetImage(highResImagePath);
                    }
                }
            }

            // Save the modified document.
            doc.Save(outputDocPath);
        }
    }
}
