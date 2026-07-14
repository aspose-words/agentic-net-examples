using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic JPEG image
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Build a Word document that contains the JPEG image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // 3. Extract JPEG images and convert each to a high‑resolution TIFF with LZW compression
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Save the image bytes to a memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Create a temporary document that contains only this image
                    Document imgDoc = new Document();
                    DocumentBuilder imgBuilder = new DocumentBuilder(imgDoc);
                    imgBuilder.InsertImage(imgStream.ToArray());

                    // Configure TIFF save options (LZW compression, 300 dpi)
                    ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                    {
                        TiffCompression = TiffCompression.Lzw,
                        Resolution = 300 // 300 dpi
                    };

                    // Save the TIFF file
                    string tiffPath = Path.Combine(artifactsDir, $"ExtractedImage_{imageIndex}.tiff");
                    imgDoc.Save(tiffPath, tiffOptions);

                    // Verify that the file was created
                    if (!File.Exists(tiffPath))
                        throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
                }

                imageIndex++;
            }
        }

        // Ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");
    }

    // Creates a 200×200 JPEG image with a blue rectangle on a white background
    private static void CreateSampleJpeg(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }

            // Explicitly save as JPEG to guarantee the correct format
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }

        // Validate that the JPEG file was created
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create JPEG file: {filePath}");
    }
}
