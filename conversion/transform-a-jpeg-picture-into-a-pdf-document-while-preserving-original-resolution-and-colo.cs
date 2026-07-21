using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string jpegPath = "sample.jpg";
        const string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        // Create a 300x300 bitmap.
        Bitmap bitmap = new Bitmap(300, 300);
        // Obtain a graphics object to draw on the bitmap.
        Graphics graphics = Graphics.FromImage(bitmap);
        // Fill the background with a solid color.
        graphics.Clear(Color.CornflowerBlue);
        // Draw a simple ellipse.
        graphics.FillEllipse(Brushes.Gold, 50, 50, 200, 200);
        // Save the bitmap as a JPEG with maximum quality (100).
        ImageCodecInfo jpegCodec = GetJpegCodec();
        EncoderParameters encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 100L);
        bitmap.Save(jpegPath, jpegCodec, encoderParams);
        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Load the JPEG into a Word document and insert it.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);

        // -----------------------------------------------------------------
        // 3. Configure PDF save options to preserve original image data.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Keep colors unchanged.
            ColorMode = ColorMode.Normal,
            // Preserve JPEG quality.
            JpegQuality = 100,
            // Do not downsample images.
            DownsampleOptions = { DownsampleImages = false }
        };
        // Use automatic image compression (keeps original JPEG bytes when possible).
        pdfOptions.ImageCompression = PdfImageCompression.Auto;

        // -----------------------------------------------------------------
        // 4. Save the document as PDF.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("PDF conversion failed: output file was not created.");
        }

        // Optional: clean up the temporary JPEG file.
        try
        {
            File.Delete(jpegPath);
        }
        catch
        {
            // Ignored – not critical for the example.
        }
    }

    // Helper method to retrieve the JPEG codec for saving images.
    private static ImageCodecInfo GetJpegCodec()
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == ImageFormat.Jpeg.Guid)
                return codec;
        }
        throw new InvalidOperationException("JPEG codec not found.");
    }
}
