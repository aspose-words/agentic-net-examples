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
        // Prepare working directories
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample TIFF image with a known DPI (150)
        string tiffPath = Path.Combine(workDir, "sample.tif");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
            }
            bmp.SetResolution(150f, 150f);
            bmp.Save(tiffPath, ImageFormat.Tiff);
        }

        // Insert the TIFF image into a Word document multiple times
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(tiffPath);
            builder.Writeln(); // separate images
        }

        // Save and reload the document to ensure proper image handling
        string docPath = Path.Combine(workDir, "sample.docx");
        doc.Save(docPath);
        Document loadedDoc = new Document(docPath);

        // Find all shapes that contain images (including the inserted TIFFs)
        var imageShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage)
                                   .ToList();

        if (!imageShapes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int index = 0;
        foreach (var shape in imageShapes)
        {
            // Export the image data to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image with Aspose.Drawing to read DPI and pixel data
                using (Bitmap sourceBmp = new Bitmap(imageStream))
                {
                    float dpiX = sourceBmp.HorizontalResolution;
                    float dpiY = sourceBmp.VerticalResolution;

                    // Create a new bitmap for PNG output, preserving size and DPI
                    using (Bitmap pngBmp = new Bitmap(sourceBmp.Width, sourceBmp.Height))
                    {
                        pngBmp.SetResolution(dpiX, dpiY);
                        using (Graphics g = Graphics.FromImage(pngBmp))
                        {
                            g.Clear(Aspose.Drawing.Color.Transparent);
                            g.DrawImage(sourceBmp, 0, 0, sourceBmp.Width, sourceBmp.Height);
                        }

                        // Save as lossless PNG
                        string pngPath = Path.Combine(outputDir, $"image_{index}.png");
                        pngBmp.Save(pngPath, ImageFormat.Png);

                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
                    }
                }
            }
            index++;
        }

        // Validate that PNG files were created
        int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG files were generated.");
    }
}
