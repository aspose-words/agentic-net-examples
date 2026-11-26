using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a sample GIF image (single frame) with width greater than 300 pixels.
        const string inputPath = "input.gif";
        const string outputPath = "resized.gif";

        int originalWidth = 500;
        int originalHeight = 200;

        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
                }
            }
            bitmap.Save(inputPath, ImageFormat.Gif);
        }

        // Verify the input file was created.
        if (!File.Exists(inputPath) || new FileInfo(inputPath).Length == 0)
            throw new InvalidOperationException("Failed to create the input GIF image.");

        // Load the GIF and resize while preserving animation frames (single frame in this example).
        using (Image originalImage = Image.FromFile(inputPath))
        {
            int width = originalImage.Width;
            int height = originalImage.Height;

            // Determine new dimensions.
            int targetWidth = width > 300 ? 300 : width;
            int targetHeight = (int)Math.Round((double)height * targetWidth / width);

            // If resizing is needed, create a new bitmap and draw the original onto it.
            if (targetWidth != width)
            {
                using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                {
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        graphics.DrawImage(originalImage, 0, 0, targetWidth, targetHeight);
                    }
                    resizedBitmap.Save(outputPath, ImageFormat.Gif);
                }
            }
            else
            {
                // No resizing needed; copy the original.
                File.Copy(inputPath, outputPath, true);
            }
        }

        // Validate that the resized GIF exists and has a non‑zero size.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("Resized GIF was not created successfully.");

        // Optional: Output result information.
        Console.WriteLine($"Input GIF:  {new FileInfo(inputPath).Length} bytes");
        Console.WriteLine($"Resized GIF: {new FileInfo(outputPath).Length} bytes");
    }
}
