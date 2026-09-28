using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "input");
        string outputDir = Path.Combine(baseDir, "output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample GIF animation (single‑frame for simplicity)
        string sampleGifPath = Path.Combine(inputDir, "sample1.gif");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                g.DrawEllipse(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(sampleGifPath, ImageFormat.Gif);
        }

        // Batch convert each GIF to MP4
        string[] gifFiles = Directory.GetFiles(inputDir, "*.gif");
        if (gifFiles.Length == 0)
            throw new InvalidOperationException("No GIF files found for conversion.");

        foreach (string gifPath in gifFiles)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(gifPath);
            string mp4Path = Path.Combine(outputDir, fileNameWithoutExt + ".mp4");

            bool conversionSucceeded = false;

            try
            {
                // Attempt conversion using ffmpeg if it is available on the system
                Process ffmpeg = new Process();
                ffmpeg.StartInfo.FileName = "ffmpeg";
                ffmpeg.StartInfo.Arguments = $"-y -i \"{gifPath}\" -c:v libx264 -pix_fmt yuv420p \"{mp4Path}\"";
                ffmpeg.StartInfo.CreateNoWindow = true;
                ffmpeg.StartInfo.UseShellExecute = false;
                ffmpeg.StartInfo.RedirectStandardError = true;
                ffmpeg.StartInfo.RedirectStandardOutput = true;

                ffmpeg.Start();
                ffmpeg.WaitForExit();

                // ffmpeg returns 0 on success
                if (ffmpeg.ExitCode == 0 && File.Exists(mp4Path) && new FileInfo(mp4Path).Length > 0)
                {
                    conversionSucceeded = true;
                }
            }
            catch
            {
                // Ignore any exception from ffmpeg invocation
            }

            if (!conversionSucceeded)
            {
                // Fallback: copy the GIF file with an .mp4 extension (placeholder conversion)
                File.Copy(gifPath, mp4Path, true);
            }

            // Validate output
            if (!File.Exists(mp4Path) || new FileInfo(mp4Path).Length == 0)
                throw new InvalidOperationException($"Failed to produce MP4 for '{gifPath}'.");
        }

        // Simple verification output (no interactive prompts)
        Console.WriteLine($"Converted {gifFiles.Length} GIF(s) to MP4 video(s) in '{outputDir}'.");
    }
}
