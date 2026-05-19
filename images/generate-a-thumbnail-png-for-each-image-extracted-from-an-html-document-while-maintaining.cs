using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D; // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create sample images using Aspose.Drawing.
        // -----------------------------------------------------------------
        string[] sampleImagePaths = CreateSampleImages(workDir);

        // -----------------------------------------------------------------
        // 2. Create a simple HTML file that references the sample images.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(workDir, "sample.html");
        CreateSampleHtml(htmlPath, sampleImagePaths);

        // -----------------------------------------------------------------
        // 3. Load the HTML document with Aspose.Words.
        // -----------------------------------------------------------------
        Document doc = new Document(htmlPath);

        // -----------------------------------------------------------------
        // 4. Extract each image, generate a thumbnail PNG while keeping aspect ratio.
        // -----------------------------------------------------------------
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .Cast<Shape>()
                        .Where(s => s.HasImage)
                        .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No images were found in the HTML document.");

        int index = 0;
        foreach (var shape in shapes)
        {
            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image with Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Determine thumbnail size (max 100x100) while preserving aspect ratio.
                    const int maxSize = 100;
                    int thumbWidth, thumbHeight;
                    if (originalBitmap.Width > originalBitmap.Height)
                    {
                        thumbWidth = maxSize;
                        thumbHeight = (int)(originalBitmap.Height * (float)maxSize / originalBitmap.Width);
                    }
                    else
                    {
                        thumbHeight = maxSize;
                        thumbWidth = (int)(originalBitmap.Width * (float)maxSize / originalBitmap.Height);
                    }

                    // Create thumbnail bitmap.
                    using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBitmap))
                        {
                            // High quality scaling.
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                        }

                        // Save thumbnail as PNG.
                        string thumbPath = Path.Combine(workDir, $"thumb_{index}.png");
                        thumbBitmap.Save(thumbPath, ImageFormat.Png);
                    }
                }
            }

            index++;
        }

        // -----------------------------------------------------------------
        // 5. Validate that thumbnails were created.
        // -----------------------------------------------------------------
        var thumbFiles = Directory.GetFiles(workDir, "thumb_*.png");
        if (thumbFiles.Length == 0)
            throw new InvalidOperationException("Thumbnail generation failed – no PNG files were created.");

        Console.WriteLine($"Generated {thumbFiles.Length} thumbnail(s) in '{workDir}'.");
    }

    // Creates two deterministic sample images and returns their file paths.
    private static string[] CreateSampleImages(string folder)
    {
        string[] paths = new string[2];

        // First image: 200x150, red background.
        using (Bitmap bmp1 = new Bitmap(200, 150))
        {
            using (Graphics g = Graphics.FromImage(bmp1))
            {
                g.Clear(Color.Red);
            }
            paths[0] = Path.Combine(folder, "sample1.png");
            bmp1.Save(paths[0], ImageFormat.Png);
        }

        // Second image: 120x240, green background.
        using (Bitmap bmp2 = new Bitmap(120, 240))
        {
            using (Graphics g = Graphics.FromImage(bmp2))
            {
                g.Clear(Color.Green);
            }
            paths[1] = Path.Combine(folder, "sample2.png");
            bmp2.Save(paths[1], ImageFormat.Png);
        }

        return paths;
    }

    // Generates a minimal HTML file that includes the provided image files.
    private static void CreateSampleHtml(string htmlPath, string[] imagePaths)
    {
        string htmlContent = "<html><body>" +
                             string.Join("", imagePaths.Select(p => $"<img src=\"{Path.GetFileName(p)}\"/>")) +
                             "</body></html>";

        File.WriteAllText(htmlPath, htmlContent);
    }
}
