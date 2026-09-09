using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Blue);
            // Draw a simple rectangle for visual content
            using (Pen pen = new Pen(Color.Yellow, 5))
            {
                g.DrawRectangle(pen, 20, 20, 160, 160);
            }
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // 2. Create a Word document and insert the JPEG image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "Document.docx");
        doc.Save(docPath);

        // 3. Extract JPEG images, convert each to grayscale BMP, and collect output file names
        var extractedBmpFiles = ExtractAndConvertImages(doc, artifactsDir);

        // 4. Store the resulting BMP files in a zip archive
        string archivePath = Path.Combine(artifactsDir, "ImagesArchive.zip");
        CreateZipArchive(extractedBmpFiles, archivePath);

        // 5. Validation
        if (!File.Exists(archivePath) || extractedBmpFiles.Count == 0)
            throw new InvalidOperationException("Archive creation failed or no BMP files were generated.");

        // Cleanup temporary BMP files (optional)
        foreach (var file in extractedBmpFiles)
            File.Delete(file);
    }

    private static System.Collections.Generic.List<string> ExtractAndConvertImages(Document doc, string outputDir)
    {
        var bmpFiles = new System.Collections.Generic.List<string>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Obtain raw image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into Aspose.Drawing.Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap bitmap = new Bitmap(ms))
            {
                // Convert to grayscale pixel by pixel
                for (int y = 0; y < bitmap.Height; y++)
                {
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        Color original = bitmap.GetPixel(x, y);
                        int gray = (original.R + original.G + original.B) / 3;
                        Color grayColor = Color.FromArgb(gray, gray, gray);
                        bitmap.SetPixel(x, y, grayColor);
                    }
                }

                // Save as BMP
                string bmpPath = Path.Combine(outputDir, $"extracted_{imageIndex}.bmp");
                bitmap.Save(bmpPath, ImageFormat.Bmp);
                bmpFiles.Add(bmpPath);
                imageIndex++;
            }
        }

        if (bmpFiles.Count == 0)
            throw new InvalidOperationException("No JPEG images were found to extract.");

        return bmpFiles;
    }

    private static void CreateZipArchive(System.Collections.Generic.List<string> files, string zipPath)
    {
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            foreach (string filePath in files)
            {
                string entryName = Path.GetFileName(filePath);
                archive.CreateEntryFromFile(filePath, entryName);
            }
        }
    }
}
