using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic BMP image to be used as sample input.
        const string sampleBmpPath = "sample.bmp";
        CreateSampleBmp(sampleBmpPath, 200, 200);

        // Step 2: Create a Word document and insert the BMP image.
        const string docPath = "input.docx";
        CreateWordDocumentWithImage(docPath, sampleBmpPath);

        // Step 3: Load the document, extract BMP images, resize them to 640x480, and save.
        const int targetWidth = 640;
        const int targetHeight = 480;
        int resizedCount = ExtractAndResizeBmpImages(docPath, targetWidth, targetHeight);

        // Validation: ensure at least one image was resized.
        if (resizedCount == 0)
            throw new InvalidOperationException("No BMP images were extracted and resized.");

        Console.WriteLine($"Successfully resized {resizedCount} BMP image(s).");
    }

    private static void CreateSampleBmp(string filePath, int width, int height)
    {
        // Create a bitmap and fill it with a solid color.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Optionally draw a simple rectangle.
        graphics.DrawRectangle(Pens.Black, 10, 10, width - 20, height - 20);
        // Save the bitmap as BMP.
        bitmap.Save(filePath, ImageFormat.Bmp);
        // Clean up.
        graphics.Dispose();
        bitmap.Dispose();

        // Verify the file was created.
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Failed to create sample BMP image.", filePath);
    }

    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the BMP image into the document.
        builder.InsertImage(imagePath);
        // Save the document.
        doc.Save(docPath);
        // Verify the file was created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create Word document.", docPath);
    }

    private static int ExtractAndResizeBmpImages(string docPath, int targetWidth, int targetHeight)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int resizedImages = 0;
        int index = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted image to a temporary BMP file.
            string extractedPath = $"extracted-{index}.bmp";
            shape.ImageData.Save(extractedPath);

            // Ensure the extracted file is a BMP (by extension, as per creation rules).
            if (Path.GetExtension(extractedPath).Equals(".bmp", StringComparison.OrdinalIgnoreCase))
            {
                // Load the extracted BMP.
                using (Bitmap original = new Bitmap(extractedPath))
                {
                    // Create a new bitmap with the target size.
                    using (Bitmap resized = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resized))
                        {
                            // Draw the original image scaled to the new dimensions.
                            graphics.DrawImage(original, 0, 0, targetWidth, targetHeight);
                        }
                        // Save the resized BMP.
                        string resizedPath = $"resized-{index}.bmp";
                        resized.Save(resizedPath, ImageFormat.Bmp);
                        // Validate the resized file exists.
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Resized image was not saved: {resizedPath}");
                        resizedImages++;
                    }
                }
            }

            // Clean up the extracted temporary file.
            if (File.Exists(extractedPath))
                File.Delete(extractedPath);

            index++;
        }

        return resizedImages;
    }
}
