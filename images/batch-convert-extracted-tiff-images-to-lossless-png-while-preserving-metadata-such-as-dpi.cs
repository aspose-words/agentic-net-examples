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
        // Prepare input and output folders.
        string inputFolder = "input_images";
        string outputFolder = "output_images";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample TIFF images with different DPI values.
        CreateSampleTiff(Path.Combine(inputFolder, "sample1.tif"), 96f);
        CreateSampleTiff(Path.Combine(inputFolder, "sample2.tif"), 300f);

        // Batch convert each TIFF to lossless PNG while preserving DPI.
        string[] tiffFiles = Directory.GetFiles(inputFolder, "*.tif");
        foreach (string tiffPath in tiffFiles)
        {
            // Load the TIFF into a temporary Word document to use Aspose.Words image handling.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(tiffPath);

            // Retrieve the inserted shape that contains the image.
            Shape shape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
            if (!shape.HasImage)
                continue;

            // Get the raw image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ImageBytes;

            // Load the image bytes into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    // Preserve original DPI.
                    float dpiX = bitmap.HorizontalResolution;
                    float dpiY = bitmap.VerticalResolution;
                    bitmap.SetResolution(dpiX, dpiY);

                    // Save as PNG with the same DPI.
                    string outputPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(tiffPath) + ".png");
                    bitmap.Save(outputPath, ImageFormat.Png);
                }
            }
        }

        // Validate that at least one PNG was created.
        int pngCount = Directory.GetFiles(outputFolder, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG files were created during conversion.");

        // Example completed successfully.
        Console.WriteLine($"Conversion completed. {pngCount} PNG file(s) saved to '{outputFolder}'.");
    }

    // Helper method to create a deterministic TIFF image with a specific DPI.
    private static void CreateSampleTiff(string filePath, float dpi)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            bitmap.SetResolution(dpi, dpi);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to have visible content.
                g.FillRectangle(Brushes.LightBlue, 20, 20, width - 40, height - 40);
            }
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}
