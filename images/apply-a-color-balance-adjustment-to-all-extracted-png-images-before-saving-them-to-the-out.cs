using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple red rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    graphics.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(sampleImagePath);
        }

        // Create a Word document and insert the sample PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Load the document (optional, already in memory) and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap bmp = new Bitmap(imageStream))
                {
                    // Apply a simple color balance adjustment.
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            Color pixel = bmp.GetPixel(x, y);
                            int r = Math.Min(255, (int)(pixel.R * 1.2)); // Increase red.
                            int g = pixel.G; // Keep green unchanged.
                            int b = Math.Max(0, (int)(pixel.B * 0.8)); // Decrease blue.
                            Color adjusted = Color.FromArgb(pixel.A, r, g, b);
                            bmp.SetPixel(x, y, adjusted);
                        }
                    }

                    // Save the adjusted image to the output folder.
                    string outputPath = Path.Combine(outputFolder, $"image-{imageIndex}.png");
                    bmp.Save(outputPath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was saved.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Optional: clean up sample files (comment out if inspection is needed).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
