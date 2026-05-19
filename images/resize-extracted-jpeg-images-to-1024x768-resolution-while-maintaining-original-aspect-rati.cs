using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for sample image and document
        string inputImagePath = Path.Combine(artifactsDir, "input.jpg");
        string docPath = Path.Combine(artifactsDir, "doc.docx");

        // Create a sample JPEG image larger than the target size (2000x1500)
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(2000, 1500))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                g.DrawString(
                    "Sample Image",
                    new Aspose.Drawing.Font("Arial", 48),
                    Aspose.Drawing.Brushes.Black,
                    new Aspose.Drawing.PointF(100, 100));
            }
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // Create a document and insert the sample image
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Extract JPEG images, resize them to fit within 1024x768 while preserving aspect ratio, and save
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Get image bytes from the shape
                byte[] imageBytes = shape.ImageData.ToByteArray();

                using (var ms = new MemoryStream(imageBytes))
                {
                    // Load the original image using Aspose.Drawing
                    using (Aspose.Drawing.Image originalImage = Aspose.Drawing.Image.FromStream(ms))
                    {
                        int origW = originalImage.Width;
                        int origH = originalImage.Height;

                        // Compute scaling factor to fit within 1024x768 without upscaling
                        double maxW = 1024;
                        double maxH = 768;
                        double factor = Math.Min(maxW / origW, maxH / origH);
                        if (factor > 1) factor = 1;

                        int newW = (int)(origW * factor);
                        int newH = (int)(origH * factor);

                        // Resize the image
                        using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(newW, newH))
                        {
                            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                            {
                                g.InterpolationMode = Aspose.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.DrawImage(originalImage, 0, 0, newW, newH);
                            }

                            string resizedPath = Path.Combine(artifactsDir, $"resized_{imageIndex}.jpg");
                            resizedBitmap.Save(resizedPath, ImageFormat.Jpeg);

                            // Validate that the file was created
                            if (!File.Exists(resizedPath))
                                throw new Exception($"Failed to save resized image: {resizedPath}");
                        }
                    }
                }

                imageIndex++;
            }
        }

        // Ensure at least one JPEG image was processed
        if (imageIndex == 0)
            throw new Exception("No JPEG images were extracted for resizing.");
    }
}
