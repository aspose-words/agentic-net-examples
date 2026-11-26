using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic JPEG sample image.
        const string sampleImagePath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle.
                using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the JPEG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document for image extraction.
        Document loadedDoc = new Document(docPath);

        // Prepare to process JPEG images.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int jpegCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure stream is at the beginning.
                using (Bitmap original = new Bitmap(ms))
                {
                    // Apply grayscale filter.
                    for (int y = 0; y < original.Height; y++)
                    {
                        for (int x = 0; x < original.Width; x++)
                        {
                            Aspose.Drawing.Color pixelColor = original.GetPixel(x, y);
                            int luminance = (int)(0.299 * pixelColor.R + 0.587 * pixelColor.G + 0.114 * pixelColor.B);
                            Aspose.Drawing.Color gray = Aspose.Drawing.Color.FromArgb(luminance, luminance, luminance);
                            original.SetPixel(x, y, gray);
                        }
                    }

                    // Save the processed image.
                    string outputImagePath = $"extracted-{jpegCount}.jpg";
                    original.Save(outputImagePath, ImageFormat.Jpeg);

                    // Validate that the file was created.
                    if (!File.Exists(outputImagePath))
                        throw new InvalidOperationException($"Failed to save processed image: {outputImagePath}");

                    jpegCount++;
                }
            }
        }

        // Ensure at least one JPEG image was processed.
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        // Optional: clean up sample files (comment out if inspection is needed).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
