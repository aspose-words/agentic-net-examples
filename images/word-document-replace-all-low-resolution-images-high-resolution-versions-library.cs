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
            // Set up temporary working folders.
            string tempRoot = Path.Combine(Path.GetTempPath(), "ImageReplacementExample");
            Directory.CreateDirectory(tempRoot);

            string sourceDocPath = Path.Combine(tempRoot, "Input.docx");
            string highResImagesFolder = Path.Combine(tempRoot, "HighResImages");
            Directory.CreateDirectory(highResImagesFolder);
            string outputDocPath = Path.Combine(tempRoot, "Output.docx");

            // Create a low‑resolution image file (1 × 1 pixel PNG).
            string lowResImageName = "sample.png";
            string lowResImagePath = Path.Combine(tempRoot, lowResImageName);
            WriteBase64ToFile(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=",
                lowResImagePath);

            // Create a high‑resolution replacement image (10 × 10 red PNG) with the same file name.
            string highResImagePath = Path.Combine(highResImagesFolder, lowResImageName);
            WriteBase64ToFile(
                "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAG0lEQVQoU2NkYGBg+M+ABYwMDAwMDAwAAAD//wMAF6cK9QAAAABJRU5ErkJggg==",
                highResImagePath);

            // Build a simple document that contains the low‑resolution image.
            Document doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.InsertImage(lowResImagePath);
            doc.Save(sourceDocPath);

            // Load the document and set the image shape's Title to the file name (used as a key).
            Document loadedDoc = new Document(sourceDocPath);
            NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.IsImage)
                {
                    shape.Title = lowResImageName; // key for replacement
                }
            }

            // Replace low‑resolution images with high‑resolution versions.
            const int lowResolutionByteThreshold = 50000; // 50 KB (our test image is far smaller)
            foreach (Shape shape in shapeNodes)
            {
                if (!shape.IsImage)
                    continue;

                if (shape.ImageData.ImageBytes.Length < lowResolutionByteThreshold)
                {
                    string imageKey = shape.Title;
                    if (string.IsNullOrEmpty(imageKey))
                        continue;

                    string candidatePath = Path.Combine(highResImagesFolder, imageKey);
                    if (File.Exists(candidatePath))
                    {
                        shape.ImageData.SetImage(candidatePath);
                    }
                }
            }

            // Save the modified document.
            loadedDoc.Save(outputDocPath);

            Console.WriteLine($"Processed document saved to: {outputDocPath}");
        }

        private static void WriteBase64ToFile(string base64, string filePath)
        {
            byte[] data = Convert.FromBase64String(base64);
            File.WriteAllBytes(filePath, data);
        }
    }
}
