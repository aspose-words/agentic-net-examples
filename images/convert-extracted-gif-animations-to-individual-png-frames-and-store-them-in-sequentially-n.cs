using System;
using System.IO;
using System.Linq;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare paths
        string gifPath = "sample.gif";
        string outputDir = "output";

        // Ensure output directory exists
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Create a sample animated GIF with three colored frames
        CreateSampleGif(gifPath);

        // Load the GIF
        using (Image gifImage = Image.FromFile(gifPath))
        {
            // Get the number of frames in the animation
            int frameCount = gifImage.GetFrameCount(FrameDimension.Time);
            if (frameCount == 0)
                throw new InvalidOperationException("No frames found in the GIF.");

            // Extract each frame and save as PNG
            for (int i = 0; i < frameCount; i++)
            {
                gifImage.SelectActiveFrame(FrameDimension.Time, i);
                using (Bitmap frameBitmap = new Bitmap(gifImage))
                {
                    string framePath = Path.Combine(outputDir, $"frame-{i}.png");
                    frameBitmap.Save(framePath, ImageFormat.Png);
                }
            }
        }

        // Validate that PNG files were created
        int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG frames were extracted.");
    }

    private static void CreateSampleGif(string outputPath)
    {
        // Define colors for frames
        Aspose.Drawing.Color[] colors = { Aspose.Drawing.Color.Red, Aspose.Drawing.Color.Green, Aspose.Drawing.Color.Blue };
        int width = 100;
        int height = 100;

        // Encoder for GIF
        ImageCodecInfo gifEncoder = ImageCodecInfo.GetImageEncoders()
            .First(enc => enc.FormatID == ImageFormat.Gif.Guid);
        Encoder saveFlagEncoder = Encoder.SaveFlag;
        EncoderParameters encoderParams = new EncoderParameters(1);

        // Create first frame
        using (Bitmap firstFrame = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(firstFrame))
            {
                g.Clear(colors[0]);
            }

            encoderParams.Param[0] = new EncoderParameter(saveFlagEncoder, (long)EncoderValue.MultiFrame);
            firstFrame.Save(outputPath, gifEncoder, encoderParams);

            // Add remaining frames
            for (int i = 1; i < colors.Length; i++)
            {
                using (Bitmap frame = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(frame))
                    {
                        g.Clear(colors[i]);
                    }

                    encoderParams.Param[0] = new EncoderParameter(saveFlagEncoder, (long)EncoderValue.FrameDimensionTime);
                    firstFrame.SaveAdd(frame, encoderParams);
                }
            }

            // Flush and close the GIF file
            encoderParams.Param[0] = new EncoderParameter(saveFlagEncoder, (long)EncoderValue.Flush);
            firstFrame.SaveAdd(encoderParams);
        }
    }
}
