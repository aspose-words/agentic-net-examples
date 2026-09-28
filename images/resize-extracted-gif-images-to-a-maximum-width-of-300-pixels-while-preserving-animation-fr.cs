using System;
using System.IO;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;   // Required for InterpolationMode

public class Program
{
    public static void Main()
    {
        // Paths for the sample input GIF and the resized output GIF.
        const string inputPath = "input.gif";
        const string outputPath = "resized.gif";
        const int maxWidth = 300;

        // Clean up any previous files.
        if (File.Exists(inputPath)) File.Delete(inputPath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // -----------------------------------------------------------------
        // Create a deterministic sample GIF (500x400) with a simple shape.
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(500, 400))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Brush brush = new SolidBrush(Color.CornflowerBlue))
                {
                    g.FillRectangle(brush, 50, 50, 400, 300);
                }
            }

            // Save the bitmap as a GIF file.
            bitmap.Save(inputPath, ImageFormat.Gif);
        }

        // --------------------------------------------------------------
        // Load the GIF, resize if necessary while preserving animation.
        // --------------------------------------------------------------
        using (Image originalImage = Image.FromFile(inputPath))
        {
            // If the image width is already within the limit, just copy it.
            if (originalImage.Width <= maxWidth)
            {
                originalImage.Save(outputPath, ImageFormat.Gif);
            }
            else
            {
                // Compute new dimensions while keeping the aspect ratio.
                double ratio = (double)maxWidth / originalImage.Width;
                int newWidth = maxWidth;
                int newHeight = (int)(originalImage.Height * ratio);

                // Create a new bitmap with the target size.
                using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                {
                    using (Graphics g = Graphics.FromImage(resizedBitmap))
                    {
                        // Use high‑quality bicubic interpolation for better results.
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                    }

                    // Save the resized bitmap as a GIF.
                    // For animated GIFs this preserves the animation frames.
                    resizedBitmap.Save(outputPath, ImageFormat.Gif);
                }
            }
        }

        // Verify that the output GIF was created successfully.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Resized GIF was not created successfully.");
        }

        Console.WriteLine($"Resized GIF saved to '{outputPath}'.");
    }
}
