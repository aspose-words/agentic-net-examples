using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace ExtractVideoFrameImages
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // -----------------------------------------------------------------
            // 1. Create a sample high‑resolution image that will act as a video frame.
            // -----------------------------------------------------------------
            string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
            using (Bitmap bitmap = new Bitmap(800, 600))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.CornflowerBlue);
                // Additional deterministic drawing can be added here if needed.
                bitmap.Save(sampleImagePath, ImageFormat.Png);
            }

            // -----------------------------------------------------------------
            // 2. Create a DOCX document and insert the sample image.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(artifactsDir, "VideoFrames.docx");
            doc.Save(docPath);

            // -----------------------------------------------------------------
            // 3. Load the document and extract all images (video frames) as PNG.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(docPath);
            NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

            int extractedCount = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Force PNG output regardless of original format.
                    string outputPath = Path.Combine(artifactsDir, $"extracted_{extractedCount}.png");

                    // Get the image bytes from the shape.
                    byte[] imageBytes = shape.ImageData.ToByteArray();

                    // Load the bytes into an Aspose.Drawing.Image and save as PNG.
                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(ms))
                    using (Bitmap bmp = new Bitmap(img))
                    {
                        bmp.Save(outputPath, ImageFormat.Png);
                    }

                    extractedCount++;
                }
            }

            // Validate that at least one image was extracted.
            if (extractedCount == 0)
                throw new InvalidOperationException("No images were extracted from the document.");
        }
    }
}
