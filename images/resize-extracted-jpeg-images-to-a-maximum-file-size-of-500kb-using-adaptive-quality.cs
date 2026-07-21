using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size for the resized JPEG images (500 KB).
    private const long MaxFileSizeBytes = 500 * 1024;

    public static void Main()
    {
        // Directories for temporary artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing.
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath, 1200, 800);

        // 2. Create a Word document and insert the sample JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleJpegPath);
        string docPath = Path.Combine(artifactsDir, "input.docx");
        doc.Save(docPath);

        // 3. Load the document (demonstrating load rule usage).
        Document loadedDoc = new Document(docPath);

        // 4. Extract JPEG images, resize adaptively to meet the size constraint.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // If already within size limit, just write it out.
                if (originalStream.Length <= MaxFileSizeBytes)
                {
                    string outPath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.jpg");
                    File.WriteAllBytes(outPath, originalStream.ToArray());
                    imageIndex++;
                    continue;
                }

                // Adaptive quality reduction loop.
                int quality = 100;
                bool saved = false;
                while (quality >= 10 && !saved)
                {
                    // Create a temporary document containing the image.
                    Document tempDoc = new Document();
                    DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                    originalStream.Position = 0;
                    tempBuilder.InsertImage(originalStream);

                    // Prepare JPEG save options with the current quality.
                    ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                    {
                        JpegQuality = quality
                    };

                    // Save to a temporary stream to check size.
                    using (MemoryStream resizedStream = new MemoryStream())
                    {
                        tempDoc.Save(resizedStream, jpegOptions);
                        if (resizedStream.Length <= MaxFileSizeBytes)
                        {
                            // Write the resized image to disk.
                            string outPath = Path.Combine(artifactsDir, $"resized_{imageIndex}.jpg");
                            File.WriteAllBytes(outPath, resizedStream.ToArray());
                            saved = true;
                        }
                        else
                        {
                            // Reduce quality and try again.
                            quality -= 10;
                        }
                    }
                }

                // If unable to meet size constraint, save the original image as fallback.
                if (!saved)
                {
                    string outPath = Path.Combine(artifactsDir, $"fallback_{imageIndex}.jpg");
                    File.WriteAllBytes(outPath, originalStream.ToArray());
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one image file was written.
        string[] outputFiles = Directory.GetFiles(artifactsDir, "*.jpg");
        if (outputFiles.Length == 0)
            throw new InvalidOperationException("No JPEG images were extracted or resized.");
    }

    // Helper method to create a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple rectangle for visual content.
            using (SolidBrush brush = new SolidBrush(Color.FromArgb(180, 70, 130, 180)))
            {
                graphics.FillRectangle(brush, 100, 100, width - 200, height - 200);
            }
            // Save as JPEG with default quality.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }
}
