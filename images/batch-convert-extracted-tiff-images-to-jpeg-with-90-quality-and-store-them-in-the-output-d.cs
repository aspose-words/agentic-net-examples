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
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputImages");
        string outputFolder = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 1: Create a sample TIFF image (deterministic content)
        // -----------------------------------------------------------------
        string sampleTiffPath = Path.Combine(baseDir, "sample.tif");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            // Save as TIFF
            bitmap.Save(sampleTiffPath, ImageFormat.Tiff);
        }

        // -----------------------------------------------------------------
        // Step 2: Insert the TIFF image into a Word document
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleTiffPath);
        string docPath = Path.Combine(baseDir, "sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Step 3: Extract all images from the document (they will be TIFF)
        // -----------------------------------------------------------------
        doc = new Document(docPath); // Reload to ensure consistency
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(inputFolder, $"image-{imageIndex}.tif");
                shape.ImageData.Save(extractedPath);
                imageIndex++;
            }
        }

        // Validate that at least one TIFF was extracted
        string[] tiffFiles = Directory.GetFiles(inputFolder, "*.tif");
        if (tiffFiles.Length == 0)
            throw new Exception("No TIFF images were extracted from the document.");

        // -----------------------------------------------------------------
        // Step 4: Batch convert extracted TIFF images to JPEG (90% quality)
        // -----------------------------------------------------------------
        // Get JPEG encoder
        ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
            .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

        // Encoder parameter for quality = 90L
        EncoderParameter qualityParam = new EncoderParameter(Encoder.Quality, 90L);
        EncoderParameters encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = qualityParam;

        foreach (string tiffPath in tiffFiles)
        {
            using (Bitmap tiffBitmap = new Bitmap(tiffPath))
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(tiffPath);
                string jpegPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.jpg");
                tiffBitmap.Save(jpegPath, jpegCodec, encoderParams);
            }
        }

        // Validate that at least one JPEG was created
        string[] jpegFiles = Directory.GetFiles(outputFolder, "*.jpg");
        if (jpegFiles.Length == 0)
            throw new Exception("No JPEG files were created during conversion.");

        // Cleanup temporary files (optional)
        if (File.Exists(sampleTiffPath)) File.Delete(sampleTiffPath);
        if (File.Exists(docPath)) File.Delete(docPath);
    }
}
