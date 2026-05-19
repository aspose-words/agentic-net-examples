using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing namespace
using Aspose.Drawing.Imaging;      // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Prepare a deterministic folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Create a source document and insert the JPEG image.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.InsertImage(jpegPath);
        string sourceDocPath = Path.Combine(artifactsDir, "source.docx");
        sourceDoc.Save(sourceDocPath);

        // 3. Load the source document and extract JPEG images.
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Save the image data to a memory stream.
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading.

                    // 4. Create a temporary document containing only this image.
                    Document tempDoc = new Document();
                    DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                    tempBuilder.InsertImage(imgStream.ToArray());

                    // 5. Configure TIFF save options with LZW compression and high resolution.
                    ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                    {
                        TiffCompression = TiffCompression.Lzw,
                        Resolution = 300 // DPI
                    };

                    // 6. Save the image as a high‑resolution TIFF.
                    string tiffPath = Path.Combine(artifactsDir, $"image_{imageIndex}.tiff");
                    tempDoc.Save(tiffPath, tiffOptions);
                    imageIndex++;
                }
            }
        }

        // Validation: ensure at least one TIFF file was created.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        Console.WriteLine($"Converted {imageIndex} image(s) to TIFF with LZW compression.");
    }

    // Creates a deterministic 200x200 JPEG image with a blue rectangle.
    private static void CreateSampleJpeg(string filePath)
    {
        // Use Aspose.Drawing.Bitmap and related types explicitly.
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (Pen pen = new Pen(Color.Blue, 5))
            {
                graphics.DrawRectangle(pen, 20, 20, 160, 160);
            }

            // Save as JPEG.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }
}
