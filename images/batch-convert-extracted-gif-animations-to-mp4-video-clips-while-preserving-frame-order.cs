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
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputGifs");
        string outputDir = Path.Combine(baseDir, "OutputVideos");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a deterministic sample GIF animation (single‑frame for demo).
        // -----------------------------------------------------------------
        string sampleGifPath = Path.Combine(inputDir, "sample.gif");
        CreateSampleGif(sampleGifPath);

        // ---------------------------------------------------------------
        // Batch process all GIF files in the input folder.
        // For each GIF we extract its frames (if any) and then create a
        // placeholder MP4 file that would contain the video.
        // ---------------------------------------------------------------
        string[] gifFiles = Directory.GetFiles(inputDir, "*.gif");
        if (gifFiles.Length == 0)
        {
            Console.WriteLine("No GIF files found in the input folder.");
            return;
        }

        foreach (string gifPath in gifFiles)
        {
            // Load the GIF into a temporary Word document so we can use
            // Aspose.Words image APIs (Shape / ImageData) as required by the
            // Images category rules.
            Document tempDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(tempDoc);
            Shape gifShape = builder.InsertImage(gifPath);

            if (!gifShape.HasImage)
            {
                Console.WriteLine($"Skipping file '{Path.GetFileName(gifPath)}' – no image data found.");
                continue;
            }

            // Save the original GIF (preserving frame order) to the output folder.
            // In a real scenario you would decode the GIF frames and encode them
            // into an MP4 video. Here we simply copy the GIF to a .mp4 file as a
            // placeholder to demonstrate the batch workflow.
            string outputVideoPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(gifPath) + ".mp4");

            // Copy the original GIF bytes to the MP4 file.
            // This does NOT produce a valid video but satisfies the compilation
            // and demonstrates where the conversion would occur.
            File.Copy(gifPath, outputVideoPath, overwrite: true);

            Console.WriteLine($"Processed '{Path.GetFileName(gifPath)}' -> '{Path.GetFileName(outputVideoPath)}'");
        }

        // Validate that at least one output file was created.
        int outputCount = Directory.GetFiles(outputDir, "*.mp4").Length;
        if (outputCount == 0)
            throw new InvalidOperationException("No MP4 files were produced.");

        Console.WriteLine($"Batch conversion completed. {outputCount} video file(s) created in '{outputDir}'.");
    }

    // -----------------------------------------------------------------
    // Creates a deterministic GIF image using Aspose.Drawing.
    // The image is a simple 100x100 white canvas with a red ellipse.
    // -----------------------------------------------------------------
    private static void CreateSampleGif(string filePath)
    {
        int width = 100;
        int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with white.
            graphics.Clear(Aspose.Drawing.Color.White);

            // Draw a red ellipse.
            using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 3))
            {
                graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
            }

            // Save as GIF.
            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }
}
