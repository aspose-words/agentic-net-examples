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
        // 1. Create a deterministic sample PNG image.
        const string samplePngPath = "sample.png";
        CreateSamplePng(samplePngPath);

        // 2. Insert the PNG into a Word document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(samplePngPath);
        doc.Save(docPath);

        // 3. Extract PNG images from the document.
        string extractedFolder = "Extracted";
        Directory.CreateDirectory(extractedFolder);
        List<string> extractedPngFiles = ExtractPngImages(doc, extractedFolder);

        if (extractedPngFiles.Count == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // 4. Convert each extracted PNG to a grayscale BMP and store in a secure folder.
        string secureArchiveFolder = "SecureArchive";
        Directory.CreateDirectory(secureArchiveFolder);
        int bmpCount = 0;

        foreach (string pngFile in extractedPngFiles)
        {
            string bmpFileName = Path.GetFileNameWithoutExtension(pngFile) + ".bmp";
            string bmpPath = Path.Combine(secureArchiveFolder, bmpFileName);
            ConvertPngToGrayscaleBmp(pngFile, bmpPath);
            bmpCount++;
        }

        if (bmpCount == 0)
            throw new InvalidOperationException("No BMP files were created during conversion.");

        Console.WriteLine($"Successfully created {bmpCount} grayscale BMP file(s) in '{secureArchiveFolder}'.");
    }

    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample PNG at '{filePath}'.");
    }

    private static List<string> ExtractPngImages(Document doc, string outputFolder)
    {
        List<string> pngFiles = new List<string>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Only process PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;

                string pngPath = Path.Combine(outputFolder,
                    $"extracted-{Guid.NewGuid()}.png");

                using (FileStream fs = new FileStream(pngPath, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(fs);
                }

                pngFiles.Add(pngPath);
            }
        }

        return pngFiles;
    }

    private static void ConvertPngToGrayscaleBmp(string pngPath, string bmpPath)
    {
        using (Bitmap original = new Bitmap(pngPath))
        {
            int width = original.Width;
            int height = original.Height;
            using (Bitmap grayBitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb))
            {
                for (int y = 0; y < height; y++)
                {
                    for (int x = 0; x < width; x++)
                    {
                        Color pixelColor = original.GetPixel(x, y);
                        // Compute luminance using standard Rec. 601 coefficients.
                        int gray = (int)(0.299 * pixelColor.R + 0.587 * pixelColor.G + 0.114 * pixelColor.B);
                        Color grayColor = Color.FromArgb(gray, gray, gray);
                        grayBitmap.SetPixel(x, y, grayColor);
                    }
                }

                grayBitmap.Save(bmpPath, ImageFormat.Bmp);
            }
        }

        if (!File.Exists(bmpPath))
            throw new InvalidOperationException($"Failed to save BMP file at '{bmpPath}'.");
    }
}
