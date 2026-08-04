using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ConvertJpegToTiff
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 200;
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");

        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (Pen pen = new Pen(Color.Blue, 5))
        {
            graphics.DrawEllipse(pen, 10, 10, imgWidth - 20, imgHeight - 20);
        }
        bitmap.Save(jpegPath, ImageFormat.Jpeg);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a source document that contains the JPEG.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.InsertImage(jpegPath);
        srcBuilder.InsertParagraph();
        srcBuilder.InsertImage(jpegPath); // insert a second copy for demonstration

        string sourceDocPath = Path.Combine(artifactsDir, "source.docx");
        sourceDoc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Load the source document and extract JPEG images.
        // -------------------------------------------------
        Document doc = new Document(sourceDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Save the extracted JPEG to a temporary file.
                string extractedJpeg = Path.Combine(artifactsDir, $"extracted_{imageIndex}.jpg");
                using (FileStream jpegStream = File.Create(extractedJpeg))
                {
                    shape.ImageData.Save(jpegStream);
                }

                // -------------------------------------------------
                // 4. Convert the extracted JPEG to high‑resolution TIFF with LZW compression.
                // -------------------------------------------------
                Document tiffDoc = new Document();
                DocumentBuilder tiffBuilder = new DocumentBuilder(tiffDoc);
                tiffBuilder.InsertImage(extractedJpeg);

                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    TiffCompression = TiffCompression.Lzw,
                    Resolution = 300 // high resolution (dpi)
                };

                string tiffPath = Path.Combine(artifactsDir, $"image_{imageIndex}.tiff");
                tiffDoc.Save(tiffPath, tiffOptions);

                // Validate that the TIFF file was created.
                if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
                    throw new Exception($"Failed to create TIFF file: {tiffPath}");

                imageIndex++;
            }
        }

        // If no JPEG images were found, raise an error.
        if (imageIndex == 0)
            throw new Exception("No JPEG images were extracted from the document.");

        // Example completed successfully.
        Console.WriteLine($"Converted {imageIndex} JPEG image(s) to TIFF with LZW compression.");
    }
}
