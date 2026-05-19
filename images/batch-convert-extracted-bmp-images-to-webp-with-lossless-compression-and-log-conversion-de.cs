using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchBmpToWebp
{
    public static void Main()
    {
        // Prepare deterministic folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchBmpToWebpDemo");
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create sample BMP images.
        CreateSampleBmp(Path.Combine(inputDir, "sample1.bmp"), 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleBmp(Path.Combine(inputDir, "sample2.bmp"), 150, 150, Aspose.Drawing.Color.LightCoral);

        // 2. Insert BMP images into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpPath in Directory.GetFiles(inputDir, "*.bmp"))
        {
            builder.InsertParagraph();
            builder.InsertImage(bmpPath);
        }

        string docPath = Path.Combine(baseDir, "Sample.docx");
        doc.Save(docPath);

        // 3. Load the document and extract all images (save them as BMP files).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int bmpIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Save the image data as a BMP file regardless of its original type.
                string extractedBmpPath = Path.Combine(inputDir, $"extracted_{bmpIndex}.bmp");
                // Ensure the image is saved in BMP format.
                using (MemoryStream ms = new MemoryStream())
                {
                    // Save the image to a memory stream first.
                    shape.ImageData.Save(ms);
                    ms.Position = 0;
                    // Load the image via Aspose.Drawing to re‑encode as BMP.
                    using (Bitmap bmp = new Bitmap(ms))
                    {
                        bmp.Save(extractedBmpPath, ImageFormat.Bmp);
                    }
                }
                bmpIndex++;
            }
        }

        // Validate that BMP files were extracted.
        string[] extractedBmpFiles = Directory.GetFiles(inputDir, "extracted_*.bmp");
        if (extractedBmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP images were extracted for conversion.");

        // 4. Convert each extracted BMP to lossless WebP.
        int webpIndex = 0;
        foreach (string bmpFile in extractedBmpFiles)
        {
            // Load BMP into a temporary document.
            Document tempDoc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
            tempBuilder.InsertImage(bmpFile);

            // Prepare WebP save options (lossless is default when no quality is set).
            ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);

            string webpPath = Path.Combine(outputDir, $"converted_{webpIndex}.webp");
            tempDoc.Save(webpPath, webpOptions);

            // Log conversion details.
            FileInfo bmpInfo = new FileInfo(bmpFile);
            FileInfo webpInfo = new FileInfo(webpPath);
            Console.WriteLine($"Converted '{Path.GetFileName(bmpFile)}' ({bmpInfo.Length} bytes) to '{Path.GetFileName(webpPath)}' ({webpInfo.Length} bytes).");

            webpIndex++;
        }

        // Validate that at least one WebP file was created.
        if (Directory.GetFiles(outputDir, "*.webp").Length == 0)
            throw new InvalidOperationException("WebP conversion failed; no output files were generated.");
    }

    // Helper method to create a deterministic BMP image.
    private static void CreateSampleBmp(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            // Draw a simple diagonal line for visual distinction.
            using (Pen pen = new Pen(Aspose.Drawing.Color.Black))
            {
                graphics.DrawLine(pen, 0, 0, width, height);
            }
            bitmap.Save(filePath, ImageFormat.Bmp);
        }
    }
}
