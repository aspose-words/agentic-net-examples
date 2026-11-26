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
        // Create a deterministic sample TIFF image.
        const string tiffPath = "sample.tiff";
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red rectangle.
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bmp.Save(tiffPath, ImageFormat.Tiff);
        }

        // Create a Word document and insert the TIFF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tiffPath);
        const string docPath = "DocumentWithTiff.docx";
        doc.Save(docPath);

        // Prepare output folder.
        const string outputFolder = "ConvertedImages";
        Directory.CreateDirectory(outputFolder);

        // Extract images from the document, convert TIFF to grayscale JPEG.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int convertedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Save the shape image to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image using Aspose.Drawing.
                using (Bitmap original = new Bitmap(imgStream))
                {
                    // Convert to grayscale.
                    using (Bitmap grayBmp = new Bitmap(original.Width, original.Height))
                    {
                        for (int y = 0; y < original.Height; y++)
                        {
                            for (int x = 0; x < original.Width; x++)
                            {
                                Aspose.Drawing.Color pixel = original.GetPixel(x, y);
                                int gray = (int)(0.3 * pixel.R + 0.59 * pixel.G + 0.11 * pixel.B);
                                Aspose.Drawing.Color grayColor = Aspose.Drawing.Color.FromArgb(gray, gray, gray);
                                grayBmp.SetPixel(x, y, grayColor);
                            }
                        }

                        // Save as JPEG.
                        string outputPath = Path.Combine(outputFolder, $"image-{convertedCount + 1}.jpg");
                        grayBmp.Save(outputPath, ImageFormat.Jpeg);
                        if (!File.Exists(outputPath))
                            throw new Exception($"Failed to create output file: {outputPath}");

                        convertedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was converted.
        if (convertedCount == 0)
            throw new Exception("No images were extracted and converted.");

        // Cleanup temporary files (optional).
        // File.Delete(tiffPath);
        // File.Delete(docPath);
    }
}
