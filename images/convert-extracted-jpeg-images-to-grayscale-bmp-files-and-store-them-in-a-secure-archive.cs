using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample JPEG image.
        const string jpegPath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                g.FillRectangle(new SolidBrush(Color.DarkRed), 50, 50, 100, 100);
            }
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the JPEG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        doc.Save(docPath);

        // Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> bmpFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Load image bytes into a bitmap.
            using (MemoryStream ms = new MemoryStream(shape.ImageData.ImageBytes))
            {
                ms.Position = 0;
                using (Bitmap original = new Bitmap(ms))
                {
                    // Create a grayscale bitmap.
                    using (Bitmap grayBitmap = new Bitmap(original.Width, original.Height))
                    {
                        for (int y = 0; y < original.Height; y++)
                        {
                            for (int x = 0; x < original.Width; x++)
                            {
                                Color pixel = original.GetPixel(x, y);
                                int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                                Color grayColor = Color.FromArgb(gray, gray, gray);
                                grayBitmap.SetPixel(x, y, grayColor);
                            }
                        }

                        // Save as BMP.
                        string bmpPath = $"image_{imageIndex}.bmp";
                        grayBitmap.Save(bmpPath, ImageFormat.Bmp);
                        bmpFiles.Add(bmpPath);
                        imageIndex++;
                    }
                }
            }
        }

        // Validate that at least one BMP file was created.
        if (bmpFiles.Count == 0)
            throw new InvalidOperationException("No JPEG images were extracted and converted.");

        // Store BMP files in a secure archive (ZIP).
        const string archivePath = "secure_archive.zip";
        using (FileStream zipToOpen = new FileStream(archivePath, FileMode.Create))
        {
            using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
            {
                foreach (string filePath in bmpFiles)
                {
                    archive.CreateEntryFromFile(filePath, Path.GetFileName(filePath));
                }
            }
        }

        // Validate that the archive was created.
        if (!File.Exists(archivePath))
            throw new InvalidOperationException("Failed to create the secure archive.");

        // Cleanup temporary files (optional).
        File.Delete(jpegPath);
        File.Delete(docPath);
        foreach (string filePath in bmpFiles)
            File.Delete(filePath);
    }
}
