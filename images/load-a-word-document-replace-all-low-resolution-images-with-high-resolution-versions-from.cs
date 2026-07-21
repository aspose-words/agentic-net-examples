using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace ImageResolutionReplacement
{
    public class Program
    {
        // Threshold for considering an image low‑resolution (in pixels).
        private const int LowResolutionPixelThreshold = 150;

        public static void Main()
        {
            // File names used in the example.
            const string lowResImagePath = "lowRes.png";
            const string highResImagePath = "highRes.png";
            const string inputDocPath = "input.docx";
            const string outputDocPath = "output.docx";

            // -----------------------------------------------------------------
            // 1. Create sample low‑resolution and high‑resolution images.
            // -----------------------------------------------------------------
            CreateSampleImage(lowResImagePath, 100, 100);   // 100 × 100 px
            CreateSampleImage(highResImagePath, 400, 400); // 400 × 400 px

            // -----------------------------------------------------------------
            // 2. Build a Word document that contains low‑resolution images.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert three low‑resolution images separated by page breaks.
            for (int i = 0; i < 3; i++)
            {
                builder.InsertImage(lowResImagePath);
                if (i < 2)
                    builder.InsertBreak(BreakType.PageBreak);
            }

            doc.Save(inputDocPath);

            // -----------------------------------------------------------------
            // 3. Load the document and replace low‑resolution images.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(inputDocPath);
            var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

            int replacedCount = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Retrieve image size information.
                ImageSize size = shape.ImageData.ImageSize;

                // Determine if the image is below the resolution threshold.
                if (size.WidthPixels < LowResolutionPixelThreshold ||
                    size.HeightPixels < LowResolutionPixelThreshold)
                {
                    // Replace the image with the high‑resolution version.
                    shape.ImageData.SetImage(highResImagePath);
                    replacedCount++;
                }
            }

            // Save the modified document.
            loadedDoc.Save(outputDocPath);

            // -----------------------------------------------------------------
            // 4. Validation.
            // -----------------------------------------------------------------
            if (!File.Exists(outputDocPath))
                throw new InvalidOperationException("The output document was not created.");

            if (replacedCount == 0)
                throw new InvalidOperationException("No low‑resolution images were found to replace.");

            // Example completed successfully.
        }

        // Creates a deterministic PNG image using Aspose.Drawing.
        private static void CreateSampleImage(string filePath, int width, int height)
        {
            // Ensure any existing file is removed.
            if (File.Exists(filePath))
                File.Delete(filePath);

            // Create a bitmap and fill it with a solid color.
            using (var bitmap = new Bitmap(width, height))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightGray);
                // Optionally, draw a simple rectangle to make the image distinguishable.
                graphics.DrawRectangle(new Pen(Color.DarkGray, 2), 0, 0, width - 1, height - 1);
                bitmap.Save(filePath);
            }
        }
    }
}
