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
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample animated GIF (2 frames) using Aspose.Drawing
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        CreateSampleAnimatedGif(gifPath);

        // 2. Insert the GIF into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // 3. Load the document and extract GIF images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                // Extract image bytes
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Load the GIF using Aspose.Drawing.Image
                    using (Image gifImage = Image.FromStream(imgStream))
                    {
                        // Preserve frame timing (delay) – read the PropertyItem 0x5100 (FrameDelay)
                        // The delay is stored in hundredths of a second per frame.
                        int[] frameDelays = null;
                        foreach (var prop in gifImage.PropertyItems)
                        {
                            if (prop.Id == 0x5100) // FrameDelay
                            {
                                int count = prop.Value.Length / 4;
                                frameDelays = new int[count];
                                for (int i = 0; i < count; i++)
                                    frameDelays[i] = BitConverter.ToInt32(prop.Value, i * 4);
                                break;
                            }
                        }

                        // Convert to animated PNG.
                        // Aspose.Drawing does not provide native APNG support, so we save each frame as a PNG
                        // and note the original delays. For demonstration, we create a single PNG from the first frame.
                        // In a real scenario, a dedicated APNG library would be required.
                        string pngPath = Path.Combine(artifactsDir, $"Gif_{gifCount}.png");
                        gifImage.SelectActiveFrame(FrameDimension.Time, 0);
                        gifImage.Save(pngPath, ImageFormat.Png);

                        // Validation
                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");

                        // Output timing information (for demonstration)
                        if (frameDelays != null && frameDelays.Length > 0)
                        {
                            Console.WriteLine($"GIF {gifCount} has {frameDelays.Length} frames. First frame delay: {frameDelays[0]} hundredths of a second.");
                        }
                        else
                        {
                            Console.WriteLine($"GIF {gifCount} has no frame delay information.");
                        }
                    }
                }

                gifCount++;
            }
        }

        // Ensure at least one GIF was processed
        if (gifCount == 0)
            throw new InvalidOperationException("No GIF images were found in the document.");

        Console.WriteLine($"Processed {gifCount} GIF image(s) and saved as PNG.");
    }

    // Helper method to create a simple animated GIF with two colored frames
    private static void CreateSampleAnimatedGif(string filePath)
    {
        // First frame (red)
        using (Bitmap bmp1 = new Bitmap(100, 100))
        using (Graphics g1 = Graphics.FromImage(bmp1))
        {
            g1.Clear(Aspose.Drawing.Color.Red);
            // Second frame (green)
            using (Bitmap bmp2 = new Bitmap(100, 100))
            using (Graphics g2 = Graphics.FromImage(bmp2))
            {
                g2.Clear(Aspose.Drawing.Color.Green);

                // Encoder parameters for GIF animation
                ImageCodecInfo gifCodec = GetEncoderInfo("image/gif");
                Encoder encoder = Encoder.SaveFlag;
                EncoderParameters encParams = new EncoderParameters(2);

                // First frame - start of the animation
                encParams.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.MultiFrame);
                encParams.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionLZW);
                bmp1.Save(filePath, gifCodec, encParams);

                // Subsequent frames
                encParams.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.FrameDimensionTime);
                bmp1.SaveAdd(bmp2, encParams);

                // Finish the animation
                encParams.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.Flush);
                bmp1.SaveAdd(encParams);
            }
        }
    }

    // Retrieves the ImageCodecInfo for the specified mime type
    private static ImageCodecInfo GetEncoderInfo(string mimeType)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.MimeType.Equals(mimeType, StringComparison.OrdinalIgnoreCase))
                return codec;
        }
        throw new InvalidOperationException($"Encoder not found for mime type {mimeType}");
    }
}
