using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;   // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);
        string imagesDir = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesDir);
        string archivePath = Path.Combine(baseDir, "GrayscaleImages.zip");

        // 1. Create a sample JPEG image.
        string jpegPath = Path.Combine(imagesDir, "sample.jpg");
        CreateSampleJpeg(jpegPath, 200, 200);

        // 2. Create a Word document and insert the JPEG image.
        string docPath = Path.Combine(baseDir, "Document.docx");
        CreateDocumentWithImage(docPath, jpegPath);

        // 3. Load the document and extract JPEG images, convert them to grayscale BMP.
        string[] bmpFiles = ExtractAndConvertImages(docPath, imagesDir);

        // 4. Validate that at least one BMP file was created.
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        // 5. Store the BMP files in a zip archive (secure archive).
        CreateZipArchive(bmpFiles, archivePath);

        // 6. Validate archive creation.
        if (!File.Exists(archivePath) || new FileInfo(archivePath).Length == 0)
            throw new InvalidOperationException("Failed to create the archive.");

        // Example completed.
    }

    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        // Create a deterministic JPEG image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.LightBlue);
            // Draw a simple rectangle.
            g.FillRectangle(new SolidBrush(Color.Crimson), 20, 20, width - 40, height - 40);
            // Save as JPEG using Aspose.Drawing.Imaging.ImageFormat.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static string[] ExtractAndConvertImages(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        var bmpList = new System.Collections.Generic.List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Set grayscale rendering for the shape's image data.
            shape.ImageData.GrayScale = true;

            // Determine output BMP file name.
            string bmpPath = Path.Combine(outputDir, $"image_{imageIndex}.bmp");

            // Save the image as BMP.
            shape.ImageData.Save(bmpPath);
            bmpList.Add(bmpPath);
            imageIndex++;
        }

        return bmpList.ToArray();
    }

    private static void CreateZipArchive(string[] files, string zipPath)
    {
        // Ensure any existing archive is removed.
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string file in files)
            {
                string entryName = Path.GetFileName(file);
                archive.CreateEntryFromFile(file, entryName);
            }
        }
    }
}
