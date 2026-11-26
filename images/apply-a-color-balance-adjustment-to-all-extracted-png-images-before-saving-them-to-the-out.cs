using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names and folders
        string inputImagePath = "sample.png";
        string docPath = "sample.docx";
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Step 1: Create a sample PNG image using Aspose.Drawing
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a Word document and insert the PNG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract PNG images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Ensure the image is a PNG
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // Reset stream position before reading

                // Load the image into Aspose.Drawing.Bitmap
                using (Bitmap originalBitmap = new Bitmap(imgStream))
                {
                    // -------------------------------------------------
                    // Step 4: Apply a simple color balance adjustment
                    // Increase the red channel by 20% (clamped to 255)
                    // -------------------------------------------------
                    Bitmap adjustedBitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height);
                    for (int y = 0; y < originalBitmap.Height; y++)
                    {
                        for (int x = 0; x < originalBitmap.Width; x++)
                        {
                            Color pixel = originalBitmap.GetPixel(x, y);
                            int newR = Math.Min(255, (int)(pixel.R * 1.2));
                            Color newColor = Color.FromArgb(pixel.A, newR, pixel.G, pixel.B);
                            adjustedBitmap.SetPixel(x, y, newColor);
                        }
                    }

                    // Save the adjusted image to the output folder
                    string outputPath = Path.Combine(outputFolder, $"extracted-{extractedCount}.png");
                    adjustedBitmap.Save(outputPath, ImageFormat.Png);
                    adjustedBitmap.Dispose();

                    extractedCount++;
                }
            }
        }

        // -------------------------------------------------
        // Validation: ensure at least one image was processed
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Cleanup temporary files (optional)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
