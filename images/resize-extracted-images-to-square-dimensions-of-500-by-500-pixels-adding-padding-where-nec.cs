using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string originalDocPath = "original.docx";

        // 1. Create a sample image (300x200) and save it locally
        CreateSampleImage(inputImagePath, 300, 200);

        // 2. Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(originalDocPath);

        // 3. Load the document and extract images
        Document loadedDoc = new Document(originalDocPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .OfType<Shape>()
                                  .Where(s => s.HasImage)
                                  .ToList();

        if (!shapeNodes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in shapeNodes)
        {
            // Save the original image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading

                // Load the image using Aspose.Drawing
                using (Image originalImage = Image.FromStream(originalStream))
                {
                    // Determine new size while preserving aspect ratio
                    const int targetSize = 500;
                    double scale = Math.Min((double)targetSize / originalImage.Width, (double)targetSize / originalImage.Height);
                    int newWidth = (int)(originalImage.Width * scale);
                    int newHeight = (int)(originalImage.Height * scale);

                    // Create a new square bitmap with white background
                    using (Bitmap squareBitmap = new Bitmap(targetSize, targetSize))
                    {
                        using (Graphics graphics = Graphics.FromImage(squareBitmap))
                        {
                            graphics.Clear(Color.White);
                            // Calculate position to center the resized image
                            int offsetX = (targetSize - newWidth) / 2;
                            int offsetY = (targetSize - newHeight) / 2;
                            // Draw the resized original image onto the square canvas
                            graphics.DrawImage(originalImage, offsetX, offsetY, newWidth, newHeight);
                        }

                        // Save the padded square image
                        string outputPath = $"resized_{imageIndex}.png";
                        squareBitmap.Save(outputPath, ImageFormat.Png);
                        imageIndex++;
                    }
                }
            }
        }

        // Validation: ensure at least one resized image was written
        if (imageIndex == 0)
            throw new InvalidOperationException("No resized images were produced.");

        // Cleanup: optional removal of temporary files (commented out)
        // File.Delete(inputImagePath);
        // File.Delete(originalDocPath);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        // Use fully qualified Aspose.Drawing types to avoid ambiguity
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, FontStyle.Bold))
                {
                    graphics.DrawString("Sample", font, Brushes.Black, new PointF(10, 10));
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
