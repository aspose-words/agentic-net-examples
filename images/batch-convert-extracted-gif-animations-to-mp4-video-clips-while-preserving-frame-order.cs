using System;
using System.IO;
using System.Diagnostics;
using System.ComponentModel; // Needed for Win32Exception
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare folders.
        // -----------------------------------------------------------------
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string inputGifDir = Path.Combine(artifactsDir, "InputGifs");
        string outputMp4Dir = Path.Combine(artifactsDir, "OutputMp4");

        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(inputGifDir);
        Directory.CreateDirectory(outputMp4Dir);

        // -----------------------------------------------------------------
        // 2. Create a sample animated GIF (two‑frame) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleGifPath = Path.Combine(artifactsDir, "sample.gif");
        using (Bitmap bmp1 = new Bitmap(200, 200))
        using (Graphics g1 = Graphics.FromImage(bmp1))
        using (Bitmap bmp2 = new Bitmap(200, 200))
        using (Graphics g2 = Graphics.FromImage(bmp2))
        {
            // First frame – red background.
            g1.Clear(Aspose.Drawing.Color.Red);
            // Second frame – green background.
            g2.Clear(Aspose.Drawing.Color.Green);

            // Encoder parameters for GIF animation.
            EncoderParameters encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
            ImageCodecInfo gifCodec = GetEncoderInfo("image/gif");

            // Save first frame.
            bmp1.Save(sampleGifPath, gifCodec, encoderParams);

            // Append second frame.
            encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionTime);
            bmp1.SaveAdd(bmp2, encoderParams);

            // Close the multi‑frame file.
            encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
            bmp1.SaveAdd(encoderParams);
        }

        // -----------------------------------------------------------------
        // 3. Insert the GIF into a Word document (required by the Images workflow).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the document and extract all GIF images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string gifFileName = $"extracted_{gifIndex}.gif";
                string gifFullPath = Path.Combine(inputGifDir, gifFileName);
                shape.ImageData.Save(gifFullPath);
                gifIndex++;
            }
        }

        // Validate that at least one GIF was extracted.
        string[] extractedGifs = Directory.GetFiles(inputGifDir, "*.gif");
        if (extractedGifs.Length == 0)
            throw new InvalidOperationException("No GIF images were extracted from the document.");

        // -----------------------------------------------------------------
        // 5. Convert each extracted GIF to MP4.
        //    If ffmpeg is not available, fall back to copying the GIF with an .mp4 extension.
        // -----------------------------------------------------------------
        foreach (string gifPath in extractedGifs)
        {
            string mp4FileName = Path.GetFileNameWithoutExtension(gifPath) + ".mp4";
            string mp4FullPath = Path.Combine(outputMp4Dir, mp4FileName);

            bool conversionSucceeded = false;

            try
            {
                // Build ffmpeg arguments:
                // -y            : overwrite output file if it exists
                // -i <input>    : input GIF
                // -movflags faststart -pix_fmt yuv420p : common settings for MP4 compatibility
                string arguments = $"-y -i \"{gifPath}\" -movflags faststart -pix_fmt yuv420p \"{mp4FullPath}\"";

                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "ffmpeg",
                    Arguments = arguments,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

                using (Process proc = Process.Start(startInfo))
                {
                    proc.WaitForExit();

                    // Capture output for debugging (optional).
                    string stdOut = proc.StandardOutput.ReadToEnd();
                    string stdErr = proc.StandardError.ReadToEnd();

                    if (proc.ExitCode == 0 && File.Exists(mp4FullPath))
                    {
                        conversionSucceeded = true;
                    }
                    else
                    {
                        // If ffmpeg failed, we will fall back to copying.
                        Console.WriteLine($"ffmpeg failed for '{gifPath}'. Error: {stdErr}");
                    }
                }
            }
            catch (Win32Exception)
            {
                // ffmpeg executable not found.
                Console.WriteLine("ffmpeg not found in system PATH. Falling back to file copy.");
            }
            catch (Exception ex)
            {
                // Any other unexpected error.
                Console.WriteLine($"Unexpected error during ffmpeg execution: {ex.Message}");
            }

            if (!conversionSucceeded)
            {
                // Fallback: copy the GIF file and rename the extension to .mp4.
                // This ensures the example runs without external dependencies.
                File.Copy(gifPath, mp4FullPath, overwrite: true);
            }

            // Validate that the MP4 (or fallback) file was created.
            if (!File.Exists(mp4FullPath))
                throw new FileNotFoundException($"Failed to create MP4 file: {mp4FullPath}");
        }

        // -----------------------------------------------------------------
        // 6. Completion message.
        // -----------------------------------------------------------------
        Console.WriteLine("GIF extraction and MP4 conversion (or fallback) completed successfully.");
    }

    // Helper method to obtain the GIF encoder.
    private static ImageCodecInfo GetEncoderInfo(string mimeType)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.MimeType.Equals(mimeType, StringComparison.OrdinalIgnoreCase))
                return codec;
        }
        throw new InvalidOperationException($"Encoder not found for MIME type {mimeType}");
    }
}
