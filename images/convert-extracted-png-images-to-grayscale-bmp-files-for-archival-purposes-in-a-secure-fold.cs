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
        // Step 1: Create a deterministic PNG image.
        const string pngPath = "input.png";
        CreateSamplePng(pngPath);

        // Step 2: Create a Word document and insert the PNG image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, pngPath);

        // Step 3: Load the document and extract PNG images.
        var document = new Document(docPath);
        var shapes = document.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        string secureFolder = "SecureArchive";
        Directory.CreateDirectory(secureFolder);

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the image to a memory stream.
            using (var imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the PNG into a bitmap.
                using (var originalBitmap = new Bitmap(imageStream))
                {
                    // Convert to grayscale.
                    using (var grayBitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height))
                    {
                        for (int y = 0; y < originalBitmap.Height; y++)
                        {
                            for (int x = 0; x < originalBitmap.Width; x++)
                            {
                                var pixel = originalBitmap.GetPixel(x, y);
                                int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                                var grayColor = Color.FromArgb(gray, gray, gray);
                                grayBitmap.SetPixel(x, y, grayColor);
                            }
                        }

                        // Save the grayscale BMP to the secure folder.
                        string outputPath = Path.Combine(secureFolder, $"image-{imageIndex}.bmp");
                        grayBitmap.Save(outputPath, ImageFormat.Bmp);

                        // Validate that the file was created.
                        if (!File.Exists(outputPath))
                            throw new Exception($"Failed to save grayscale BMP: {outputPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new Exception("No PNG images were found and converted.");
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;

        // Create bitmap and draw a red ellipse on white background.
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        var pen = new Pen(Color.Red, 5);
        graphics.DrawEllipse(pen, 20, 20, width - 40, height - 40);

        // Save as PNG.
        bitmap.Save(filePath, ImageFormat.Png);

        // Cleanup drawing resources.
        pen.Dispose();
        graphics.Dispose();
        bitmap.Dispose();

        // Validate that the PNG was created.
        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample PNG image: {filePath}");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Validate that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create Word document: {docPath}");
    }
}
