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
        // Create a sample JPEG image (2000x1500) using Aspose.Drawing.
        const string sampleImagePath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(2000, 1500))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to have some content.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 100, 100, 1800, 1300);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Create a new Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "document.docx";
        doc.Save(docPath);

        // Reload the document to simulate extraction scenario.
        Document loadedDoc = new Document(docPath);
        var imageShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                                   .ToList();

        if (!imageShapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int index = 0;
        foreach (var shape in imageShapes)
        {
            // Extract the image bytes into a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image using Aspose.Drawing.
                using (Image originalImage = Image.FromStream(imageStream))
                {
                    // Determine scaling factor to fit within 1024x768 while preserving aspect ratio.
                    double maxWidth = 1024.0;
                    double maxHeight = 768.0;
                    double widthRatio = maxWidth / originalImage.Width;
                    double heightRatio = maxHeight / originalImage.Height;
                    double scale = Math.Min(widthRatio, heightRatio);
                    // If the image is already smaller than the target size, keep original dimensions.
                    if (scale > 1.0) scale = 1.0;

                    int newWidth = (int)(originalImage.Width * scale);
                    int newHeight = (int)(originalImage.Height * scale);

                    // Resize the image.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized JPEG image.
                        string outputPath = $"resized_{index}.jpg";
                        resizedBitmap.Save(outputPath, ImageFormat.Jpeg);
                    }
                }
            }

            index++;
        }

        // Validate that at least one resized image file was created.
        if (index == 0)
            throw new InvalidOperationException("No resized images were saved.");

        // Cleanup temporary files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
