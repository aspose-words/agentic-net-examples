using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageInversionExample
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Draw a solid red rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    graphics.FillRectangle(brush, 10, 10, 80, 80);
                }
            }
            // Save the sample image as PNG.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Step 2: Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "DocumentWithImages.docx";
        doc.Save(docPath);

        // Step 3: Load the document (already in memory) and extract PNG images.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Save the original extracted PNG.
                string extractedPath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(extractedPath);

                // Load the extracted image for processing.
                using (Bitmap bitmap = new Bitmap(extractedPath))
                {
                    // Invert colors pixel by pixel.
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            Color original = bitmap.GetPixel(x, y);
                            Color inverted = Color.FromArgb(255 - original.R, 255 - original.G, 255 - original.B);
                            bitmap.SetPixel(x, y, inverted);
                        }
                    }

                    // Save the inverted image.
                    string invertedPath = $"inverted-{extractedCount}.png";
                    bitmap.Save(invertedPath, ImageFormat.Png);

                    // Validate that the inverted image was saved.
                    if (!File.Exists(invertedPath))
                        throw new Exception($"Inverted image was not saved: {invertedPath}");
                }

                extractedCount++;
            }
        }

        // Validate that at least one PNG image was processed.
        if (extractedCount == 0)
            throw new Exception("No PNG images were found and processed in the document.");

        Console.WriteLine($"Processed {extractedCount} PNG image(s). Inverted images saved with prefix 'inverted-'.");
    }
}
