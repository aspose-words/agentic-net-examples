using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare sample JPEG images.
        string[] jpegFiles = { "sample1.jpg", "sample2.jpg" };
        CreateSampleJpeg(jpegFiles[0], 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleJpeg(jpegFiles[1], 150, 150, Aspose.Drawing.Color.LightCoral);

        // Create a Word document and insert the JPEG images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string jpegPath in jpegFiles)
        {
            builder.InsertImage(jpegPath);
            builder.Writeln(); // separate images
        }
        string docPath = "DocumentWithImages.docx";
        doc.Save(docPath);

        // Load the document for extraction.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> savedBmpFiles = new List<string>();
        string archiveFolder = "SecureArchive";
        Directory.CreateDirectory(archiveFolder);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage) continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(imageStream))
                {
                    // Convert to grayscale.
                    using (Aspose.Drawing.Bitmap grayBitmap = new Aspose.Drawing.Bitmap(originalBitmap.Width, originalBitmap.Height))
                    {
                        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(grayBitmap))
                        {
                            for (int y = 0; y < originalBitmap.Height; y++)
                            {
                                for (int x = 0; x < originalBitmap.Width; x++)
                                {
                                    Aspose.Drawing.Color pixelColor = originalBitmap.GetPixel(x, y);
                                    int luminance = (int)(0.299 * pixelColor.R + 0.587 * pixelColor.G + 0.114 * pixelColor.B);
                                    Aspose.Drawing.Color gray = Aspose.Drawing.Color.FromArgb(luminance, luminance, luminance);
                                    grayBitmap.SetPixel(x, y, gray);
                                }
                            }
                        }

                        string bmpPath = Path.Combine(archiveFolder, $"image-{imageIndex}.bmp");
                        grayBitmap.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
                        savedBmpFiles.Add(bmpPath);
                        imageIndex++;
                    }
                }
            }
        }

        // Validation: ensure at least one BMP file was created.
        if (savedBmpFiles.Count == 0)
            throw new InvalidOperationException("No JPEG images were extracted and converted.");

        foreach (string bmpFile in savedBmpFiles)
        {
            if (!File.Exists(bmpFile))
                throw new FileNotFoundException($"Expected BMP file was not created: {bmpFile}");
        }

        // Cleanup sample JPEG files and the intermediate document.
        foreach (string jpegPath in jpegFiles)
        {
            if (File.Exists(jpegPath))
                File.Delete(jpegPath);
        }
        if (File.Exists(docPath))
            File.Delete(docPath);
    }

    private static void CreateSampleJpeg(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backgroundColor);
                // Draw a simple rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }
}
