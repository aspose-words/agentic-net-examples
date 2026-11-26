using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Base working directory
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(workDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files that contain the sample image
        int docCount = 3;
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(workDir, $"Document{i}.docx");
            CreateDocWithImage(docPath, sampleImagePath);
            docPaths.Add(docPath);
        }

        // Prepare folders for extracted images and thumbnails
        string extractedDir = Path.Combine(workDir, "Extracted");
        string thumbDir = Path.Combine(workDir, "Thumbnails");
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(thumbDir);

        // Batch extract images and create thumbnails
        int extractedCount = 0;
        foreach (string docPath in docPaths)
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imgIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    imgIndex++;
                    // Save original image
                    string imgFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Img{imgIndex}.png";
                    string imgFullPath = Path.Combine(extractedDir, imgFileName);
                    using (MemoryStream imgStream = new MemoryStream(shape.ImageData.ImageBytes))
                    {
                        imgStream.Position = 0;
                        using (Bitmap bmp = new Bitmap(imgStream))
                        {
                            bmp.Save(imgFullPath, ImageFormat.Png);
                        }
                    }

                    // Create thumbnail (100x100)
                    string thumbPath = Path.Combine(thumbDir, $"Thumb_{imgFileName}");
                    CreateThumbnail(imgFullPath, thumbPath, 100, 100);
                    extractedCount++;
                }
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Create PDF catalog with thumbnails
        string catalogPath = Path.Combine(workDir, "Catalog.pdf");
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);
        builder.PageSetup.PaperSize = PaperSize.A4;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        string[] thumbFiles = Directory.GetFiles(thumbDir, "*.png");
        foreach (string thumbFile in thumbFiles)
        {
            builder.Writeln(Path.GetFileNameWithoutExtension(thumbFile));
            builder.InsertImage(thumbFile);
            builder.Writeln(); // Add spacing
        }

        catalog.Save(catalogPath, SaveFormat.Pdf);

        if (!File.Exists(catalogPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");

        // Cleanup (optional)
        // Console.WriteLine("Processing completed successfully.");
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
                using (Brush brush = new SolidBrush(Aspose.Drawing.Color.LightGray))
                {
                    g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void CreateDocWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Docx);
    }

    private static void CreateThumbnail(string sourcePath, string thumbPath, int thumbWidth, int thumbHeight)
    {
        using (Bitmap sourceBmp = new Bitmap(sourcePath))
        {
            using (Bitmap thumbBmp = new Bitmap(thumbWidth, thumbHeight))
            {
                using (Graphics g = Graphics.FromImage(thumbBmp))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(sourceBmp, new Rectangle(0, 0, thumbWidth, thumbHeight));
                }
                thumbBmp.Save(thumbPath, ImageFormat.Png);
            }
        }
    }
}
