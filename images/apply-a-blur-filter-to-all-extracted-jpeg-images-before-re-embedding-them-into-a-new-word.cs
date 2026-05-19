using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create sample JPEG images using Aspose.Drawing
        // -----------------------------------------------------------------
        string[] sampleImageFiles = { Path.Combine(artifactsDir, "sample1.jpg"),
                                      Path.Combine(artifactsDir, "sample2.jpg") };

        for (int i = 0; i < sampleImageFiles.Length; i++)
        {
            using (Bitmap bmp = new Bitmap(200, 200))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                // Draw a simple colored ellipse to make the image recognizable
                using (SolidBrush brush = new SolidBrush(i % 2 == 0 ? Color.Red : Color.Blue))
                {
                    g.FillEllipse(brush, 20, 20, 160, 160);
                }
                // Save as JPEG
                using (FileStream fs = new FileStream(sampleImageFiles[i], FileMode.Create, FileAccess.Write))
                {
                    bmp.Save(fs, ImageFormat.Jpeg);
                }
            }
        }

        // -----------------------------------------------------------------
        // 2. Create a source document and insert the sample JPEG images
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        foreach (string imgPath in sampleImageFiles)
        {
            builder.InsertParagraph();
            builder.InsertImage(imgPath);
        }

        string sourceDocPath = Path.Combine(artifactsDir, "Source.docx");
        sourceDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 3. Load the document, extract JPEG images, apply a simple blur,
        //    and replace them in the document
        // -----------------------------------------------------------------
        Document doc = new Document(sourceDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int jpegCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Get original image bytes
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Load original image into Aspose.Drawing.Bitmap
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            using (Bitmap originalBitmap = new Bitmap(originalStream))
            {
                // Create a new bitmap for the blurred image
                using (Bitmap blurredBitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height))
                using (Graphics g = Graphics.FromImage(blurredBitmap))
                {
                    // Simple blur approximation: draw the image scaled down then up
                    double scale = 0.9; // 90% size
                    int w = (int)(originalBitmap.Width * scale);
                    int h = (int)(originalBitmap.Height * scale);
                    int x = (originalBitmap.Width - w) / 2;
                    int y = (originalBitmap.Height - h) / 2;

                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(originalBitmap, new Rectangle(x, y, w, h));

                    // Save blurred bitmap to a memory stream as JPEG
                    using (MemoryStream blurredStream = new MemoryStream())
                    {
                        blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                        blurredStream.Position = 0; // Reset before reuse

                        // Replace the image in the shape with the blurred version
                        shape.ImageData.SetImage(blurredStream);
                    }
                }
            }

            jpegCount++;
        }

        // Validation: ensure at least one JPEG image was processed
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found to process.");

        // -----------------------------------------------------------------
        // 4. Save the modified document
        // -----------------------------------------------------------------
        string outputDocPath = Path.Combine(artifactsDir, "Output.docx");
        doc.Save(outputDocPath, SaveFormat.Docx);
    }
}
