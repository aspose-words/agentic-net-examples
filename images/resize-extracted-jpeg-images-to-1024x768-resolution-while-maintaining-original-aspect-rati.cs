using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageResizeExample
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image (2000x1500) and save it locally.
        const string inputImagePath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(2000, 1500))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle to have some content.
                graphics.FillRectangle(new SolidBrush(Color.LightBlue), 100, 100, 1800, 1300);
            }
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Only process JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image using Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Determine new size while preserving aspect ratio within 1024x768.
                    const int maxWidth = 1024;
                    const int maxHeight = 768;
                    double widthRatio = (double)maxWidth / originalBitmap.Width;
                    double heightRatio = (double)maxHeight / originalBitmap.Height;
                    double scale = Math.Min(widthRatio, heightRatio);
                    if (scale > 1) // Image already smaller than target size.
                        scale = 1;

                    int newWidth = (int)(originalBitmap.Width * scale);
                    int newHeight = (int)(originalBitmap.Height * scale);

                    // Resize the image.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image.
                        string outputImagePath = $"resized-{imageIndex}.jpg";
                        resizedBitmap.Save(outputImagePath, ImageFormat.Jpeg);
                        Console.WriteLine($"Resized image saved to: {outputImagePath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found and resized in the document.");
    }
}
