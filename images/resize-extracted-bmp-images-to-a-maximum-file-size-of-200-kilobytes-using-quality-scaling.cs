using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size in bytes (200 KB)
    private const long MaxFileSize = 200 * 1024;

    public static void Main()
    {
        // Prepare deterministic folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample BMP image (800x600) and save it locally
        string sampleBmpPath = Path.Combine(artifactsDir, "sample.bmp");
        CreateSampleBmp(sampleBmpPath, 800, 600);

        // 2. Create a Word document and insert the BMP image
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleBmpPath);

        // 3. Load the document and extract images
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue; // Ensure shape actually contains image data

            // Extract original image into a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Resize/compress until the size is <= 200 KB
                using (MemoryStream resizedStream = ResizeBmpToTargetSize(originalStream, MaxFileSize))
                {
                    // Save the final BMP to a deterministic file name
                    string outputPath = Path.Combine(artifactsDir, $"output_{imageIndex}.bmp");
                    File.WriteAllBytes(outputPath, resizedStream.ToArray());

                    // Validation
                    FileInfo info = new FileInfo(outputPath);
                    if (info.Length > MaxFileSize)
                        throw new InvalidOperationException($"Resized image exceeds target size: {info.Length} bytes.");

                    imageIndex++;
                }
            }
        }

        // Ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Creates a simple BMP file with a solid color background using Aspose.Drawing
    private static void CreateSampleBmp(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            bitmap.Save(filePath, ImageFormat.Bmp);
        }
    }

    // Inserts the provided image file into a new Word document
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Resizes the image by scaling down until it fits within the target size.
    // The method always returns a BMP image stream.
    private static MemoryStream ResizeBmpToTargetSize(Stream sourceStream, long targetSize)
    {
        // Load the original bitmap from the source stream
        using (Bitmap original = new Bitmap(sourceStream))
        {
            // If the original already satisfies the size requirement, return it as‑is (as BMP)
            using (MemoryStream testStream = new MemoryStream())
            {
                original.Save(testStream, ImageFormat.Bmp);
                if (testStream.Length <= targetSize)
                {
                    testStream.Position = 0;
                    MemoryStream result = new MemoryStream();
                    testStream.CopyTo(result);
                    result.Position = 0;
                    return result;
                }
            }

            // Iteratively scale down by 90% until the size constraint is met
            Bitmap current = (Bitmap)original.Clone();
            try
            {
                while (true)
                {
                    int newWidth = Math.Max(1, (int)(current.Width * 0.9));
                    int newHeight = Math.Max(1, (int)(current.Height * 0.9));

                    using (Bitmap scaled = new Bitmap(newWidth, newHeight))
                    using (Graphics g = Graphics.FromImage(scaled))
                    {
                        g.DrawImage(current, 0, 0, newWidth, newHeight);
                        using (MemoryStream ms = new MemoryStream())
                        {
                            scaled.Save(ms, ImageFormat.Bmp);
                            if (ms.Length <= targetSize)
                            {
                                ms.Position = 0;
                                MemoryStream finalStream = new MemoryStream();
                                ms.CopyTo(finalStream);
                                finalStream.Position = 0;
                                return finalStream;
                            }
                        }
                    }

                    // Prepare for next iteration
                    Bitmap next = new Bitmap(newWidth, newHeight);
                    using (Graphics g = Graphics.FromImage(next))
                    {
                        g.DrawImage(current, 0, 0, newWidth, newHeight);
                    }
                    current.Dispose();
                    current = next;
                }
            }
            finally
            {
                current.Dispose();
            }
        }
    }
}
