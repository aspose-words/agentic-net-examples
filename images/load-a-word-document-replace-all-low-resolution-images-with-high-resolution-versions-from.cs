using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, Pen

namespace ImageResolutionReplacement
{
    public class Program
    {
        // Threshold for considering an image low‑resolution (pixel width).
        private const int LowResolutionWidthThreshold = 200;

        public static void Main()
        {
            // Prepare folders.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create sample low‑resolution and high‑resolution images.
            string lowResImagePath = Path.Combine(artifactsDir, "low.png");
            string highResImagePath = Path.Combine(artifactsDir, "high.png");
            CreateSampleImage(lowResImagePath, 100, 100);   // 100 × 100 px
            CreateSampleImage(highResImagePath, 300, 300); // 300 × 300 px

            // Build a document that contains low‑resolution images.
            string inputDocPath = Path.Combine(artifactsDir, "input.docx");
            CreateDocumentWithLowResImages(inputDocPath, lowResImagePath);

            // Load the document and replace low‑resolution images.
            string outputDocPath = Path.Combine(artifactsDir, "output.docx");
            ReplaceLowResolutionImages(inputDocPath, outputDocPath, highResImagePath);

            // Simple validation – ensure the output file exists.
            if (!File.Exists(outputDocPath))
                throw new InvalidOperationException("The output document was not created.");

            Console.WriteLine("Image replacement completed successfully.");
        }

        // Creates a deterministic PNG image with the given dimensions.
        private static void CreateSampleImage(string filePath, int width, int height)
        {
            // Use Aspose.Drawing.Bitmap and related types.
            using (var bitmap = new Bitmap(width, height))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle border to make the image visible.
                using (var pen = new Pen(Color.Black))
                {
                    graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
                bitmap.Save(filePath);
            }
        }

        // Generates a Word document that contains only low‑resolution images.
        private static void CreateDocumentWithLowResImages(string docPath, string lowResImagePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert the low‑resolution image twice to have multiple candidates.
            builder.InsertImage(lowResImagePath);
            builder.InsertParagraph(); // separate the images
            builder.InsertImage(lowResImagePath);

            doc.Save(docPath);
        }

        // Loads a document, finds images whose pixel width is below the threshold,
        // and replaces them with the high‑resolution version.
        private static void ReplaceLowResolutionImages(string inputPath, string outputPath, string highResImagePath)
        {
            var doc = new Document(inputPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true);

            int replacedCount = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Retrieve image size information.
                ImageSize size = shape.ImageData.ImageSize;

                // Determine if the image is low‑resolution based on pixel width.
                if (size.WidthPixels < LowResolutionWidthThreshold)
                {
                    // Replace the image data with the high‑resolution image.
                    shape.ImageData.SetImage(highResImagePath);
                    replacedCount++;
                }
            }

            // Ensure at least one image was replaced; otherwise, something went wrong.
            if (replacedCount == 0)
                throw new InvalidOperationException("No low‑resolution images were found to replace.");

            doc.Save(outputPath);
        }
    }
}
