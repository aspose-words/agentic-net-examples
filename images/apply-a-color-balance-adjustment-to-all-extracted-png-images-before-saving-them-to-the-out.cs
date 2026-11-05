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
        // Prepare folders
        string inputImagePath = "sample.png";
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual content
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document (simulating typical load workflow)
        Document loadedDoc = new Document(docPath);

        // Extract PNG images, apply color balance, and save them
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save the original image to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading

                // Load the image into Aspose.Drawing.Bitmap
                using (Bitmap img = new Bitmap(ms))
                {
                    // Apply a simple color balance adjustment
                    for (int y = 0; y < img.Height; y++)
                    {
                        for (int x = 0; x < img.Width; x++)
                        {
                            Color pixel = img.GetPixel(x, y);
                            int r = Math.Min(255, pixel.R + 30);
                            int g = Math.Max(0, pixel.G - 20);
                            int b = Math.Min(255, pixel.B + 10);
                            Color newColor = Color.FromArgb(pixel.A, r, g, b);
                            img.SetPixel(x, y, newColor);
                        }
                    }

                    // Save the adjusted image to the output folder
                    string outputPath = Path.Combine(outputFolder, $"image-{imageIndex}.png");
                    img.Save(outputPath, ImageFormat.Png);
                    imageIndex++;
                }
            }
        }

        // Validation: ensure at least one image was saved
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Cleanup temporary files (optional)
        File.Delete(inputImagePath);
        File.Delete(docPath);
    }
}
