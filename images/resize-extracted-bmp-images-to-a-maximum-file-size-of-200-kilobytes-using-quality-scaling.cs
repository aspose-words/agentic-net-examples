using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size: 200 KB
    private const long MaxFileSizeBytes = 200 * 1024;

    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample BMP image (800x800, solid blue)
        string sampleBmpPath = Path.Combine(artifactsDir, "sample.bmp");
        CreateSampleBmp(sampleBmpPath, 800, 800);

        // 2. Insert the BMP into a Word document
        string docPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImage(docPath, sampleBmpPath);

        // 3. Load the document and extract images
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        int resizedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get original image bytes (any format)
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Resize the image until it fits the size constraint
            byte[] resizedBytes = ResizeImageToMaxSize(originalBytes, MaxFileSizeBytes);

            // Save the resized image as BMP
            string resizedPath = Path.Combine(artifactsDir, $"resized_{imageIndex}.bmp");
            File.WriteAllBytes(resizedPath, resizedBytes);

            // Validation
            FileInfo info = new FileInfo(resizedPath);
            if (!info.Exists)
                throw new InvalidOperationException($"Resized file not created: {resizedPath}");
            if (info.Length > MaxFileSizeBytes)
                throw new InvalidOperationException($"Resized file exceeds size limit: {resizedPath}");

            resizedCount++;
            imageIndex++;
        }

        if (resizedCount == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        Console.WriteLine($"Successfully resized {resizedCount} image(s).");
    }

    // Creates a solid‑color BMP file using Aspose.Drawing
    private static void CreateSampleBmp(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.Blue);
            bitmap.Save(filePath, ImageFormat.Bmp);
        }
    }

    // Creates a Word document and inserts the specified image
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Resizes an image (any format) iteratively until its size is <= maxSizeBytes.
    // The result is always saved as BMP.
    private static byte[] ResizeImageToMaxSize(byte[] imageBytes, long maxSizeBytes)
    {
        using (MemoryStream inputStream = new MemoryStream(imageBytes))
        using (Bitmap original = new Bitmap(inputStream))
        {
            int currentWidth = original.Width;
            int currentHeight = original.Height;
            Bitmap workingBitmap = (Bitmap)original.Clone();

            while (true)
            {
                // Save current bitmap as BMP to a memory stream and check size
                using (MemoryStream ms = new MemoryStream())
                {
                    workingBitmap.Save(ms, ImageFormat.Bmp);
                    if (ms.Length <= maxSizeBytes)
                        return ms.ToArray();
                }

                // Reduce dimensions by 10%
                currentWidth = (int)(currentWidth * 0.9);
                currentHeight = (int)(currentHeight * 0.9);

                if (currentWidth < 1 || currentHeight < 1)
                    throw new InvalidOperationException("Unable to reduce image below the required size.");

                // Create a new resized bitmap
                Bitmap resized = new Bitmap(currentWidth, currentHeight);
                using (Graphics g = Graphics.FromImage(resized))
                {
                    g.DrawImage(workingBitmap, 0, 0, currentWidth, currentHeight);
                }

                // Dispose previous bitmap and continue
                workingBitmap.Dispose();
                workingBitmap = resized;
            }
        }
    }
}
