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
        // Create a deterministic sample JPEG image using Aspose.Drawing.
        string jpegPath = "sample.jpg";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(jpegPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Insert the JPEG image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int tiffCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Process only JPEG images.
                if (shape.ImageData.ImageType == ImageType.Jpeg)
                {
                    // Load the JPEG image using Aspose.Drawing.
                    using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(imageStream))
                    {
                        // Create a high‑resolution bitmap (300 DPI) and draw the original image onto it.
                        using (Aspose.Drawing.Bitmap highResBmp = new Aspose.Drawing.Bitmap(img.Width, img.Height))
                        {
                            highResBmp.SetResolution(300, 300);
                            using (Aspose.Drawing.Graphics g2 = Aspose.Drawing.Graphics.FromImage(highResBmp))
                            {
                                g2.DrawImage(img, 0, 0, img.Width, img.Height);
                            }

                            // Prepare TIFF encoder with LZW compression.
                            ImageCodecInfo tiffCodec = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == Aspose.Drawing.Imaging.ImageFormat.Tiff.Guid);
                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = new EncoderParameter(
                                Encoder.Compression,
                                (long)EncoderValue.CompressionLZW);

                            // Save the high‑resolution image as TIFF.
                            string tiffPath = $"image-{tiffCount}.tiff";
                            highResBmp.Save(tiffPath, tiffCodec, encoderParams);
                            tiffCount++;
                        }
                    }
                }
            }
        }

        // Validate that at least one TIFF file was created.
        if (tiffCount == 0)
            throw new InvalidOperationException("No JPEG images were found to convert to TIFF.");

        // Cleanup temporary files.
        File.Delete(jpegPath);
        File.Delete(docPath);
    }
}
