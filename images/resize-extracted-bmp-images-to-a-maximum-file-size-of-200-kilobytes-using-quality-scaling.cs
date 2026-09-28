using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size: 200 KB
    private const long MaxFileSize = 200 * 1024;

    public static void Main()
    {
        // Step 1: Create a deterministic BMP image.
        const string sampleBmpPath = "sample.bmp";
        CreateSampleBmp(sampleBmpPath, 800, 600);

        // Step 2: Create a Word document and insert the BMP image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, sampleBmpPath);

        // Step 3: Load the document and process images.
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Resize the bitmap until it meets the size requirement.
                    using (Bitmap resizedBitmap = ResizeBitmapToMaxSize(originalBitmap, MaxFileSize))
                    {
                        // Save the resized image as BMP.
                        string outputPath = $"resized-{imageIndex}.bmp";
                        resizedBitmap.Save(outputPath, ImageFormat.Bmp);

                        // Validate output.
                        FileInfo info = new FileInfo(outputPath);
                        if (!info.Exists || info.Length > MaxFileSize)
                        {
                            throw new InvalidOperationException($"Resized image '{outputPath}' was not created correctly.");
                        }
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were found in the document.");
        }
    }

    // Creates a deterministic BMP file with simple content.
    private static void CreateSampleBmp(string path, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
        {
            g.DrawRectangle(pen, 50, 50, width - 100, height - 100);
        }
        bitmap.Save(path, ImageFormat.Bmp);
        g.Dispose();
        bitmap.Dispose();

        // Validate creation.
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample BMP at '{path}'.");
    }

    // Inserts the given image into a new Word document.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Validate creation.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create document at '{docPath}'.");
    }

    // Resizes the bitmap iteratively until its BMP representation fits within maxSize bytes.
    private static Bitmap ResizeBitmapToMaxSize(Bitmap original, long maxSize)
    {
        int currentWidth = original.Width;
        int currentHeight = original.Height;

        // Clone the original to work on.
        Bitmap workingBitmap = (Bitmap)original.Clone();

        try
        {
            while (true)
            {
                // Check current size.
                using (MemoryStream ms = new MemoryStream())
                {
                    workingBitmap.Save(ms, ImageFormat.Bmp);
                    if (ms.Length <= maxSize)
                        break;
                }

                // Reduce dimensions by 10%.
                currentWidth = (int)(currentWidth * 0.9);
                currentHeight = (int)(currentHeight * 0.9);
                if (currentWidth < 1 || currentHeight < 1)
                    throw new InvalidOperationException("Cannot reduce image size further to meet the file size constraint.");

                // Create a new resized bitmap.
                Bitmap resized = new Bitmap(currentWidth, currentHeight);
                using (Graphics g = Graphics.FromImage(resized))
                {
                    g.DrawImage(workingBitmap, 0, 0, currentWidth, currentHeight);
                }

                // Replace the working bitmap.
                workingBitmap.Dispose();
                workingBitmap = resized;
            }

            return workingBitmap;
        }
        catch
        {
            workingBitmap.Dispose();
            throw;
        }
    }
}
