using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D; // For InterpolationMode

public class Program
{
    // Maximum allowed file size in bytes (500 KB)
    private const long MaxFileSize = 500 * 1024;

    public static void Main()
    {
        // Step 1: Create a deterministic sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath, 800, 600);

        // Step 2: Create a Word document and insert the JPEG image.
        const string inputDocPath = "input.docx";
        CreateDocumentWithImage(inputDocPath, sampleImagePath);

        // Step 3: Load the document and extract JPEG images.
        Document doc = new Document(inputDocPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(originalStream))
                {
                    // Adaptive resizing to meet the max file size.
                    Bitmap resizedBitmap = AdaptiveResize(bitmap, MaxFileSize);

                    // Save the resized image.
                    string outputImagePath = $"resized-{imageIndex}.jpg";
                    using (FileStream outStream = new FileStream(outputImagePath, FileMode.Create, FileAccess.Write))
                    {
                        resizedBitmap.Save(outStream, ImageFormat.Jpeg);
                    }

                    // Validate that the file was created and meets size requirement.
                    FileInfo info = new FileInfo(outputImagePath);
                    if (!info.Exists || info.Length > MaxFileSize)
                        throw new InvalidOperationException($"Failed to create resized image within size limit: {outputImagePath}");

                    resizedBitmap.Dispose();
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to resize.");

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete(inputDocPath);
    }

    // Creates a deterministic sample JPEG image.
    private static void CreateSampleJpeg(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, width - 100, height - 100);
                }
            }

            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                bitmap.Save(fs, ImageFormat.Jpeg);
            }
        }
    }

    // Creates a Word document and inserts the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Performs adaptive resizing by reducing dimensions until the file size is within the limit.
    private static Bitmap AdaptiveResize(Bitmap original, long maxSize)
    {
        Bitmap current = (Bitmap)original.Clone();

        while (true)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                current.Save(ms, ImageFormat.Jpeg);
                if (ms.Length <= maxSize)
                    break;
            }

            // Reduce dimensions by 10%.
            int newWidth = (int)(current.Width * 0.9);
            int newHeight = (int)(current.Height * 0.9);
            if (newWidth < 1 || newHeight < 1)
                break; // Cannot reduce further.

            Bitmap resized = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(resized))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(current, 0, 0, newWidth, newHeight);
            }

            current.Dispose();
            current = resized;
        }

        return current;
    }
}
