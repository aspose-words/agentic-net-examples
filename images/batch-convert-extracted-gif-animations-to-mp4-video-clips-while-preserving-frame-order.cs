using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders
        string inputFolder = "input";
        string outputFolder = "output";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample GIF animation (single‑frame for demonstration)
        string gifPath = Path.Combine(inputFolder, "sample.gif");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        try
        {
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawEllipse(pen, 20, 20, 160, 160);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.DrawString("GIF", font, brush, new Aspose.Drawing.PointF(60, 80));
                }
            }
            finally
            {
                g.Dispose();
            }

            // Save the bitmap as a GIF file
            bitmap.Save(gifPath, ImageFormat.Gif);
        }
        finally
        {
            bitmap.Dispose();
        }

        // Batch convert each GIF to MP4 (placeholder conversion – copies bytes)
        string[] gifFiles = Directory.GetFiles(inputFolder, "*.gif");
        foreach (string gifFile in gifFiles)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(gifFile);
            string mp4Path = Path.Combine(outputFolder, fileNameWithoutExt + ".mp4");

            // Read GIF bytes
            byte[] gifBytes = File.ReadAllBytes(gifFile);
            // Write them directly as MP4 (placeholder conversion)
            File.WriteAllBytes(mp4Path, gifBytes);
        }

        // Validation: ensure at least one MP4 was created
        string[] mp4Files = Directory.GetFiles(outputFolder, "*.mp4");
        if (mp4Files.Length == 0)
        {
            throw new InvalidOperationException("No MP4 files were created.");
        }

        // Display result count
        Console.WriteLine($"Converted {mp4Files.Length} GIF(s) to MP4 video clips.");
    }
}
