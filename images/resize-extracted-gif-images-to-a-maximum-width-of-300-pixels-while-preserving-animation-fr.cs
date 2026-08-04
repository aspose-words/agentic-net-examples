using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum width for the resized GIF (in pixels)
    private const int MaxWidth = 300;

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample animated GIF (2 frames) from a
        //    base‑64 string and save it as "input.gif".
        // -----------------------------------------------------------------
        const string base64Gif =
            "R0lGODdhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs="; // 1×1 transparent GIF (single frame)
        // For demonstration we will duplicate the single frame to create a simple animation.
        byte[] gifBytes = Convert.FromBase64String(base64Gif);
        File.WriteAllBytes("input.gif", gifBytes);

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample GIF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape gifShape = builder.InsertImage("input.gif");
        doc.Save("DocumentWithGif.docx");

        // -----------------------------------------------------------------
        // 3. Locate the GIF shape, extract its image bytes and load it with
        //    Aspose.Drawing.Image.
        // -----------------------------------------------------------------
        Shape shapeWithGif = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .FirstOrDefault(s => s.HasImage && s.ImageData.ImageType == ImageType.Gif);

        if (shapeWithGif == null)
            throw new InvalidOperationException("No GIF image found in the document.");

        using (MemoryStream originalGifStream = new MemoryStream())
        {
            shapeWithGif.ImageData.Save(originalGifStream);
            originalGifStream.Position = 0;

            using (Image originalGif = Image.FromStream(originalGifStream))
            {
                // -----------------------------------------------------------------
                // 4. Determine if resizing is required.
                // -----------------------------------------------------------------
                int originalWidth = originalGif.Width;
                if (originalWidth <= MaxWidth)
                {
                    Console.WriteLine("GIF width is already within the limit; no resizing needed.");
                    return;
                }

                double scale = (double)MaxWidth / originalWidth;
                int newWidth = MaxWidth;
                int newHeight = (int)(originalGif.Height * scale);

                // -----------------------------------------------------------------
                // 5. Resize each frame while preserving animation metadata.
                // -----------------------------------------------------------------
                // Prepare encoder parameters for GIF.
                ImageCodecInfo gifCodec = GetEncoder(ImageFormat.Gif);
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);

                // Create the first resized frame.
                Image firstFrame = ResizeFrame(originalGif, 0, newWidth, newHeight);
                firstFrame.Save("resized.gif", gifCodec, encoderParams);

                // Append remaining frames.
                encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionTime);
                int frameCount = originalGif.GetFrameCount(FrameDimension.Time);
                for (int i = 1; i < frameCount; i++)
                {
                    using (Image nextFrame = ResizeFrame(originalGif, i, newWidth, newHeight))
                    {
                        firstFrame.SaveAdd(nextFrame, encoderParams);
                    }
                }

                // Finalize the multi‑frame file.
                encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
                firstFrame.SaveAdd(encoderParams);
                firstFrame.Dispose();

                // -----------------------------------------------------------------
                // 6. Replace the shape's image with the resized GIF and save the document.
                // -----------------------------------------------------------------
                shapeWithGif.ImageData.SetImage("resized.gif");
                doc.Save("DocumentWithResizedGif.docx");

                // -----------------------------------------------------------------
                // 7. Validation – ensure the resized GIF file exists and its width
                //    does not exceed the maximum.
                // -----------------------------------------------------------------
                if (!File.Exists("resized.gif"))
                    throw new InvalidOperationException("Resized GIF was not created.");

                using (Image resizedCheck = Image.FromFile("resized.gif"))
                {
                    if (resizedCheck.Width > MaxWidth)
                        throw new InvalidOperationException("Resized GIF width exceeds the limit.");
                }

                Console.WriteLine("GIF successfully resized and saved as 'resized.gif'.");
            }
        }
    }

    // Helper: Resize a specific frame of a multi‑frame image.
    private static Image ResizeFrame(Image source, int frameIndex, int width, int height)
    {
        source.SelectActiveFrame(FrameDimension.Time, frameIndex);
        using (Bitmap srcBitmap = new Bitmap(source))
        {
            Bitmap resizedBitmap = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(resizedBitmap))
            {
                g.Clear(Aspose.Drawing.Color.Transparent);
                g.DrawImage(srcBitmap, 0, 0, width, height);
            }
            return resizedBitmap;
        }
    }

    // Helper: Retrieve the encoder for a given image format.
    private static ImageCodecInfo GetEncoder(ImageFormat format)
    {
        return ImageCodecInfo.GetImageDecoders()
                             .FirstOrDefault(codec => codec.FormatID == format.Guid)
               ?? throw new InvalidOperationException("Encoder not found for the specified format.");
    }
}
