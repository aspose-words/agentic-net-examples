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
    public static void Main()
    {
        // Prepare deterministic folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputGifPath = Path.Combine(artifactsDir, "sample.gif");
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        string outputGifPattern = Path.Combine(artifactsDir, "resized_{0}.gif");

        // -----------------------------------------------------------------
        // 1. Create a deterministic animated GIF (2 frames) from a base64 string.
        // -----------------------------------------------------------------
        const string base64Gif =
            "R0lGODdhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==";
        byte[] gifBytes = Convert.FromBase64String(base64Gif);
        File.WriteAllBytes(inputGifPath, gifBytes);

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the GIF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // InsertImage already appends the shape to the document, no extra AppendChild needed.
        Shape gifShape = builder.InsertImage(inputGifPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document, find GIF images, resize them to max 300 px width,
        //    and save the resized GIFs while preserving animation frames.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Gif) continue;

            // Extract original GIF bytes.
            byte[] originalGif = shape.ImageData.ToByteArray();

            using (MemoryStream originalStream = new MemoryStream(originalGif))
            using (Aspose.Drawing.Image originalImage = Aspose.Drawing.Image.FromStream(originalStream))
            {
                int originalWidth = originalImage.Width;
                if (originalWidth <= 300)
                {
                    // No resizing needed – just save the original GIF.
                    string outPath = string.Format(outputGifPattern, gifIndex);
                    File.WriteAllBytes(outPath, originalGif);
                    gifIndex++;
                    continue;
                }

                // Calculate new dimensions while preserving aspect ratio.
                int newWidth = 300;
                int newHeight = (int)Math.Round(originalImage.Height * (300.0 / originalWidth));

                // Encoder for GIF.
                ImageCodecInfo gifCodec = GetEncoderInfo("image/gif");
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);

                // First frame.
                using (Bitmap firstFrame = new Bitmap(newWidth, newHeight))
                using (Graphics g = Graphics.FromImage(firstFrame))
                {
                    g.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                    using (MemoryStream outStream = new MemoryStream())
                    {
                        firstFrame.Save(outStream, gifCodec, encoderParams);

                        // Subsequent frames.
                        encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionTime);
                        FrameDimension dimension = new FrameDimension(originalImage.FrameDimensionsList[0]);
                        int frameCount = originalImage.GetFrameCount(dimension);

                        for (int i = 1; i < frameCount; i++)
                        {
                            originalImage.SelectActiveFrame(dimension, i);
                            using (Bitmap resizedFrame = new Bitmap(newWidth, newHeight))
                            using (Graphics g2 = Graphics.FromImage(resizedFrame))
                            {
                                g2.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                                firstFrame.SaveAdd(resizedFrame, encoderParams);
                            }
                        }

                        // Finalize.
                        encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
                        firstFrame.SaveAdd(encoderParams);

                        // Write the resized animated GIF to disk.
                        string outPath = string.Format(outputGifPattern, gifIndex);
                        File.WriteAllBytes(outPath, outStream.ToArray());
                        gifIndex++;
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one resized GIF was produced.
        // -----------------------------------------------------------------
        if (gifIndex == 0)
            throw new InvalidOperationException("No GIF images were found or resized.");
    }

    // Helper to obtain the GIF encoder.
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
