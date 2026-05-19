using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace AsposeWordsImageMotionBlur
{
    public class Program
    {
        // Paths used in the example
        private const string ArtifactsDir = "Artifacts";
        private const string SampleImagePath = "sample.jpg";
        private const string OutputDocPath = "Artifacts/DocumentWithBlurredImages.docx";

        public static void Main()
        {
            // Ensure output folder exists
            Directory.CreateDirectory(ArtifactsDir);

            // 1. Create a deterministic sample JPEG image
            CreateSampleJpeg(SampleImagePath);

            // 2. Create a document and insert the sample JPEG several times
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Document with original JPEG images:");
            for (int i = 0; i < 3; i++)
            {
                builder.InsertImage(SampleImagePath);
                builder.Writeln(); // separate images
            }

            // 3. Extract JPEG images, apply a simple motion‑blur effect, and re‑embed them
            int processedCount = 0;
            var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;
                if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

                // Save the original image to a memory stream
                using (MemoryStream originalStream = new MemoryStream())
                {
                    shape.ImageData.Save(originalStream);
                    originalStream.Position = 0;

                    // Load the image into a bitmap
                    using (Bitmap sourceBitmap = new Bitmap(originalStream))
                    {
                        // Apply a simple horizontal motion blur
                        Bitmap blurredBitmap = ApplyMotionBlur(sourceBitmap, blurLength: 10, offsetStep: 1f);

                        // Save the blurred bitmap back to a stream (JPEG format)
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // Replace the shape's image with the blurred version
                            shape.ImageData.SetImage(blurredStream);
                            processedCount++;
                        }

                        blurredBitmap.Dispose();
                    }
                }
            }

            // Validate that at least one JPEG image was processed
            if (processedCount == 0)
                throw new InvalidOperationException("No JPEG images were found to process.");

            // 4. Save the resulting document
            doc.Save(OutputDocPath);
        }

        // Creates a simple deterministic JPEG image using Aspose.Drawing
        private static void CreateSampleJpeg(string filePath)
        {
            const int width = 200;
            const int height = 150;
            using (Bitmap bitmap = new Bitmap(width, height))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, width - 40, height - 40);
                }
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    graphics.FillRectangle(brush, 60, 60, 80, 30);
                }
                bitmap.Save(filePath, ImageFormat.Jpeg);
            }
        }

        // Generates a horizontal motion blur by drawing the source image multiple times with incremental offsets
        private static Bitmap ApplyMotionBlur(Bitmap source, int blurLength, float offsetStep)
        {
            Bitmap blurred = new Bitmap(source.Width, source.Height);
            using (Graphics g = Graphics.FromImage(blurred))
            {
                // Start with a transparent background
                g.Clear(Color.Transparent);

                // Draw the source image repeatedly, shifting it horizontally each time
                for (int i = 0; i < blurLength; i++)
                {
                    float offsetX = i * offsetStep;
                    g.DrawImage(source, offsetX, 0, source.Width, source.Height);
                }
            }
            return blurred;
        }
    }
}
