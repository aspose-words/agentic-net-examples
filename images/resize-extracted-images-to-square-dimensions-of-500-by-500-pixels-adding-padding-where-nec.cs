using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;   // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string docPath = "Original.docx";
        const string outputImagePath = "resized_0.png";

        // 1. Create a sample image (300x200) using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(300, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle for visual reference
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, 280, 180);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // 2. Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // 3. Load the document and extract the image from the shape
        Document loadedDoc = new Document(docPath);
        Shape imageShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                    .OfType<Shape>()
                                    .FirstOrDefault(s => s.HasImage);
        if (imageShape == null)
            throw new InvalidOperationException("No image found in the document.");

        // Save the image data to a memory stream
        using (MemoryStream imageStream = new MemoryStream())
        {
            imageShape.ImageData.Save(imageStream);
            imageStream.Position = 0; // Reset stream position before reading

            // 4. Load the extracted image into a Bitmap
            using (Bitmap originalBitmap = new Bitmap(imageStream))
            {
                const int targetSize = 500;
                double scale = Math.Min((double)targetSize / originalBitmap.Width,
                                        (double)targetSize / originalBitmap.Height);
                int scaledWidth = (int)Math.Round(originalBitmap.Width * scale);
                int scaledHeight = (int)Math.Round(originalBitmap.Height * scale);

                // Resize the original image while preserving aspect ratio
                using (Bitmap resizedBitmap = new Bitmap(scaledWidth, scaledHeight))
                {
                    using (Graphics gResize = Graphics.FromImage(resizedBitmap))
                    {
                        gResize.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        gResize.DrawImage(originalBitmap, 0, 0, scaledWidth, scaledHeight);
                    }

                    // Create a 500x500 canvas with white background and center the resized image
                    using (Bitmap finalBitmap = new Bitmap(targetSize, targetSize))
                    {
                        using (Graphics gCanvas = Graphics.FromImage(finalBitmap))
                        {
                            gCanvas.Clear(Color.White);
                            int offsetX = (targetSize - scaledWidth) / 2;
                            int offsetY = (targetSize - scaledHeight) / 2;
                            gCanvas.DrawImage(resizedBitmap, offsetX, offsetY, scaledWidth, scaledHeight);
                        }

                        // Save the final padded image
                        finalBitmap.Save(outputImagePath, ImageFormat.Png);
                    }
                }
            }
        }

        // 5. Validate that the output file was created
        if (!File.Exists(outputImagePath))
            throw new FileNotFoundException("Resized image was not saved.", outputImagePath);
    }
}
