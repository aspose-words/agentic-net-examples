using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare directories for artifacts and the secure archive.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string archiveDir = Path.Combine(artifactsDir, "SecureArchive");
        Directory.CreateDirectory(archiveDir);

        // 1. Create a deterministic sample JPEG image using Aspose.Drawing.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Insert the JPEG image into a new Word document.
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        CreateDocumentWithImage(jpegPath, docPath);

        // 3. Load the document, extract JPEG images, convert each to grayscale BMP, and save them.
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Apply grayscale rendering.
                shape.ImageData.GrayScale = true;

                // Save the image as BMP; the file extension determines the format.
                string bmpFileName = Path.Combine(archiveDir, $"Image_{extractedCount}.bmp");
                shape.ImageData.Save(bmpFileName);
                extractedCount++;
            }
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted from the document.");

        // 4. Create a ZIP archive containing the grayscale BMP files.
        string zipPath = Path.Combine(artifactsDir, "GrayscaleImages.zip");
        if (File.Exists(zipPath))
            File.Delete(zipPath);
        ZipFile.CreateFromDirectory(archiveDir, zipPath);

        // Validate that the archive was created successfully.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the archive.");

        // Optional cleanup: uncomment to remove temporary files.
        // Directory.Delete(archiveDir, true);
    }

    // Creates a 200x200 JPEG image with a solid background and a white rectangle.
    private static void CreateSampleJpeg(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.CornflowerBlue);
            using (Pen pen = new Pen(Aspose.Drawing.Color.White, 5))
            {
                graphics.DrawRectangle(pen, 20, 20, 160, 160);
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a new Word document and inserts the specified image.
    private static void CreateDocumentWithImage(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
