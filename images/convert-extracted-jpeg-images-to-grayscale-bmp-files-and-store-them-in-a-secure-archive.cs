using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image.
        const string jpegPath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
            }
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Step 2: Insert the JPEG image into a Word document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        doc.Save(docPath);

        // Step 3: Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> bmpFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;

            // Load JPEG into a bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap jpegBitmap = new Bitmap(ms))
                {
                    // Convert to grayscale.
                    for (int y = 0; y < jpegBitmap.Height; y++)
                    {
                        for (int x = 0; x < jpegBitmap.Width; x++)
                        {
                            Color original = jpegBitmap.GetPixel(x, y);
                            int gray = (int)(0.3 * original.R + 0.59 * original.G + 0.11 * original.B);
                            Color grayColor = Color.FromArgb(gray, gray, gray);
                            jpegBitmap.SetPixel(x, y, grayColor);
                        }
                    }

                    // Save as BMP.
                    string bmpPath = $"image-{imageIndex}.bmp";
                    jpegBitmap.Save(bmpPath, ImageFormat.Bmp);
                    if (!File.Exists(bmpPath))
                        throw new InvalidOperationException($"Failed to create BMP file: {bmpPath}");

                    bmpFiles.Add(bmpPath);
                    imageIndex++;
                }
            }
        }

        if (bmpFiles.Count == 0)
            throw new InvalidOperationException("No JPEG images were extracted from the document.");

        // Step 4: Store the BMP files in a secure archive (ZIP file).
        const string archivePath = "archive.zip";
        using (FileStream zipToOpen = new FileStream(archivePath, FileMode.Create))
        {
            using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
            {
                foreach (string bmpFile in bmpFiles)
                {
                    archive.CreateEntryFromFile(bmpFile, Path.GetFileName(bmpFile));
                }
            }
        }

        if (!File.Exists(archivePath))
            throw new InvalidOperationException($"Failed to create archive: {archivePath}");

        // Cleanup temporary files (optional).
        File.Delete(jpegPath);
        File.Delete(docPath);
        foreach (string bmpFile in bmpFiles)
        {
            File.Delete(bmpFile);
        }
    }
}
