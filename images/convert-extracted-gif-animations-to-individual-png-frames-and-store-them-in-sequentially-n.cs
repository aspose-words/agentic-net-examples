using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = "output";
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Create a sample GIF (single-frame for simplicity)
        string gifPath = "sample.gif";
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Blue);
            }
            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // Load the GIF
        using (Image gifImage = Image.FromFile(gifPath))
        {
            // Determine the number of frames in the animation (time dimension)
            int frameCount = gifImage.GetFrameCount(FrameDimension.Time);
            if (frameCount == 0)
                throw new InvalidOperationException("No frames found in the GIF.");

            // Extract each frame to a PNG file
            for (int i = 0; i < frameCount; i++)
            {
                gifImage.SelectActiveFrame(FrameDimension.Time, i);
                using (Bitmap frameBmp = new Bitmap(gifImage))
                {
                    string outPath = Path.Combine(outputDir, $"frame_{i + 1}.png");
                    frameBmp.Save(outPath, ImageFormat.Png);
                }
            }
        }

        // Validate that PNG files were created
        int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG frames were extracted.");

        // Optional: indicate success (no interactive prompts)
        Console.WriteLine($"Extracted {pngCount} PNG frame(s) to '{outputDir}'.");
    }
}
