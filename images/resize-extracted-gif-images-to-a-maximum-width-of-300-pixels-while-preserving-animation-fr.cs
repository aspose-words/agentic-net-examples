using System;
using System.IO;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a sample GIF image (static for simplicity)
        const string inputPath = "input.gif";
        const string outputPath = "output.gif";

        // Ensure any previous files are removed
        if (File.Exists(inputPath)) File.Delete(inputPath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // Create a simple bitmap and save as GIF
        using (Bitmap bitmap = new Bitmap(500, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Blue), 0, 0, 250, 200);
                g.FillRectangle(new SolidBrush(Color.Red), 250, 0, 250, 200);
            }
            bitmap.Save(inputPath, ImageFormat.Gif);
        }

        // Load the GIF
        using (Image gifImage = Image.FromFile(inputPath))
        {
            // Determine new width (max 300)
            int originalWidth = gifImage.Width;
            int originalHeight = gifImage.Height;
            int newWidth = originalWidth > 300 ? 300 : originalWidth;
            int newHeight = (int)((double)newWidth / originalWidth * originalHeight);

            // Resize while preserving animation frames (single frame case)
            using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
            {
                using (Graphics g = Graphics.FromImage(resizedBitmap))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(gifImage, 0, 0, newWidth, newHeight);
                }

                // Save the resized image as GIF
                resizedBitmap.Save(outputPath, ImageFormat.Gif);
            }
        }

        // Validate output
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new Exception("Resized GIF was not created successfully.");
        }

        // Indicate success (no interactive input)
        Console.WriteLine("Resized GIF saved to " + Path.GetFullPath(outputPath));
    }
}
