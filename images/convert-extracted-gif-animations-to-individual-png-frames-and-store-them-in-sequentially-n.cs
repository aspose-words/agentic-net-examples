using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample multi‑frame GIF image.
        const string gifPath = "sample.gif";
        CreateSampleGif(gifPath);

        // Insert the GIF into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save("doc_with_gif.docx");

        // Locate the shape that contains the GIF.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        Shape gifShape = null;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                gifShape = shape;
                break;
            }
        }

        if (gifShape == null)
            throw new Exception("No GIF image found in the document.");

        // Get the raw GIF bytes.
        byte[] gifBytes = gifShape.ImageData.ToByteArray();

        // Load the GIF with Aspose.Drawing and extract each frame as PNG.
        using (MemoryStream ms = new MemoryStream(gifBytes))
        {
            ms.Position = 0;
            using (Aspose.Drawing.Image gifImage = Aspose.Drawing.Image.FromStream(ms))
            {
                // The first entry in FrameDimensionsList corresponds to the time dimension for animated GIFs.
                FrameDimension dimension = new FrameDimension(gifImage.FrameDimensionsList[0]);
                int frameCount = gifImage.GetFrameCount(dimension);
                if (frameCount == 0)
                    throw new Exception("The GIF contains no frames.");

                for (int i = 0; i < frameCount; i++)
                {
                    gifImage.SelectActiveFrame(dimension, i);
                    using (Bitmap frameBitmap = new Bitmap(gifImage))
                    {
                        string pngPath = $"frame_{i}.png";
                        frameBitmap.Save(pngPath, ImageFormat.Png);
                    }
                }
            }
        }

        // Verify that PNG files were created.
        string[] pngFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "frame_*.png");
        if (pngFiles.Length == 0)
            throw new Exception("No PNG frames were created.");

        Console.WriteLine($"Extracted {pngFiles.Length} PNG frame(s) from the GIF.");
    }

    // Creates a 3‑frame GIF with solid red, green, and blue squares.
    private static void CreateSampleGif(string filePath)
    {
        const int width = 100;
        const int height = 100;

        using (Bitmap bmp1 = new Bitmap(width, height))
        using (Graphics g1 = Graphics.FromImage(bmp1))
        using (Bitmap bmp2 = new Bitmap(width, height))
        using (Graphics g2 = Graphics.FromImage(bmp2))
        using (Bitmap bmp3 = new Bitmap(width, height))
        using (Graphics g3 = Graphics.FromImage(bmp3))
        {
            g1.Clear(Aspose.Drawing.Color.Red);
            g2.Clear(Aspose.Drawing.Color.Green);
            g3.Clear(Aspose.Drawing.Color.Blue);

            // Encoder for saving multi‑frame GIF.
            Encoder saveFlag = Encoder.SaveFlag;
            ImageCodecInfo gifCodec = GetEncoderInfo("image/gif");

            // Save first frame.
            EncoderParameters epFirst = new EncoderParameters(1);
            epFirst.Param[0] = new EncoderParameter(saveFlag, (long)EncoderValue.MultiFrame);
            bmp1.Save(filePath, gifCodec, epFirst);

            // Add second frame.
            EncoderParameters epAdd = new EncoderParameters(1);
            epAdd.Param[0] = new EncoderParameter(saveFlag, (long)EncoderValue.FrameDimensionTime);
            bmp1.SaveAdd(bmp2, epAdd);

            // Add third frame.
            bmp1.SaveAdd(bmp3, epAdd);

            // Flush and close the multi‑frame file.
            EncoderParameters epFlush = new EncoderParameters(1);
            epFlush.Param[0] = new EncoderParameter(saveFlag, (long)EncoderValue.Flush);
            bmp1.SaveAdd(epFlush);
        }
    }

    // Retrieves the codec info for a given MIME type.
    private static ImageCodecInfo GetEncoderInfo(string mimeType)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.MimeType.Equals(mimeType, StringComparison.OrdinalIgnoreCase))
                return codec;
        }
        throw new Exception($"Encoder for MIME type '{mimeType}' not found.");
    }
}
