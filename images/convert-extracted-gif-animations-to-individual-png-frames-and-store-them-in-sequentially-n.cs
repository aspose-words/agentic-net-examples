using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string gifPath = "sample.gif";
        const string outputPattern = "frame_{0}.png";

        // -------------------------------------------------
        // Step 1: Create a deterministic sample GIF image
        // -------------------------------------------------
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red ellipse
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
                }
            }

            // Save as GIF (single‑frame animation)
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // Ensure the GIF file was created
        if (!File.Exists(gifPath))
            throw new FileNotFoundException("Failed to create the sample GIF.", gifPath);

        // -------------------------------------------------
        // Step 2: Load the GIF and extract frames as PNGs
        // -------------------------------------------------
        using (Image gifImage = Image.FromFile(gifPath))
        {
            // Determine the number of frames in the GIF (time dimension)
            int frameCount = gifImage.GetFrameCount(FrameDimension.Time);

            // If the GIF has no frames, treat it as a single image
            if (frameCount == 0)
                frameCount = 1;

            for (int i = 0; i < frameCount; i++)
            {
                // Select the current frame
                gifImage.SelectActiveFrame(FrameDimension.Time, i);

                // Create a bitmap from the current frame
                using (Bitmap frameBitmap = new Bitmap(gifImage))
                {
                    string outputPath = string.Format(outputPattern, i);
                    frameBitmap.Save(outputPath, ImageFormat.Png);
                }
            }
        }

        // -------------------------------------------------
        // Step 3: Validate that at least one PNG was created
        // -------------------------------------------------
        string[] pngFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "frame_*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG frames were extracted from the GIF.");

        // Optional: Output result count (no interactive prompt)
        Console.WriteLine($"Extracted {pngFiles.Length} PNG frame(s) from the GIF.");
    }
}
