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
        // Step 1: Create a sample JPEG image.
        const string jpegPath = "sample.jpg";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        graphics.DrawRectangle(Aspose.Drawing.Pens.Black, 20, 20, 160, 160);
        bitmap.Save(jpegPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        graphics.Dispose();
        bitmap.Dispose();

        // Step 2: Insert the JPEG into a Word document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        doc.Save(docPath);

        // Step 3: Extract JPEG images from the document and convert to high‑resolution TIFF with LZW compression.
        const string tiffPath = "converted.tiff";
        bool tiffCreated = false;

        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(imageStream))
                {
                    // Obtain the TIFF codec.
                    var tiffCodec = Aspose.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == Aspose.Drawing.Imaging.ImageFormat.Tiff.Guid);
                    if (tiffCodec == null)
                        throw new Exception("TIFF codec not found.");

                    // Prepare encoder parameters for LZW compression.
                    EncoderParameters encoderParams = new EncoderParameters(1);
                    EncoderParameter compressionParam = new EncoderParameter(
                        Encoder.Compression,
                        (long)EncoderValue.CompressionLZW);
                    encoderParams.Param[0] = compressionParam;

                    // Save as TIFF with LZW compression.
                    img.Save(tiffPath, tiffCodec, encoderParams);
                    tiffCreated = true;
                }
            }
        }

        // Validation: ensure the TIFF file was created.
        if (!tiffCreated || !File.Exists(tiffPath))
            throw new Exception("TIFF conversion failed or no image was processed.");

        // Optional cleanup (commented out).
        // File.Delete(jpegPath);
        // File.Delete(docPath);
    }
}
