using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(baseDir);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files that contain the sample image
        int docCount = 3;
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(baseDir, $"doc{i}.docx");
            CreateDocWithImage(docPath, sampleImagePath);
            docPaths.Add(docPath);
        }

        // Prepare folders for extracted images and thumbnails
        string extractedDir = Path.Combine(baseDir, "Extracted");
        string thumbDir = Path.Combine(baseDir, "Thumbnails");
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(thumbDir);

        // Extract images from each DOCX and create thumbnails
        int totalImages = 0;
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
                    totalImages++;

                    // Save original extracted image
                    string extractedPath = Path.Combine(
                        extractedDir,
                        $"{Path.GetFileNameWithoutExtension(docPath)}_img{imgIndex}.png");
                    shape.ImageData.Save(extractedPath);

                    // Create thumbnail (100x100) from extracted image
                    string thumbPath = Path.Combine(
                        thumbDir,
                        $"{Path.GetFileNameWithoutExtension(docPath)}_thumb{imgIndex}.png");
                    CreateThumbnail(extractedPath, thumbPath, 100, 100);
                }
            }
        }

        if (totalImages == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Build PDF catalog with thumbnails
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);
        string[] thumbFiles = Directory.GetFiles(thumbDir, "*.png");
        if (thumbFiles.Length == 0)
            throw new InvalidOperationException("No thumbnails were created for the PDF catalog.");

        foreach (string thumbFile in thumbFiles)
        {
            builder.InsertParagraph();
            builder.InsertImage(thumbFile);
            builder.Writeln(Path.GetFileName(thumbFile));
        }

        string pdfPath = Path.Combine(baseDir, "Catalog.pdf");
        catalog.Save(pdfPath, SaveFormat.Pdf);

        // Validate final PDF exists
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        // Create a deterministic bitmap using Aspose.Drawing
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
            }
            finally
            {
                g.Dispose();
            }

            bitmap.Save(path, ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }
    }

    private static void CreateDocWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document generated at {DateTime.Now}");
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        // No explicit Dispose needed for DocumentBuilder or Document
    }

    private static void CreateThumbnail(string sourcePath, string thumbPath, int thumbWidth, int thumbHeight)
    {
        // Load source image
        Aspose.Drawing.Bitmap sourceBitmap = new Aspose.Drawing.Bitmap(sourcePath);
        try
        {
            Aspose.Drawing.Bitmap thumbBitmap = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight);
            try
            {
                Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbBitmap);
                try
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.DrawImage(sourceBitmap, new Rectangle(0, 0, thumbWidth, thumbHeight));
                }
                finally
                {
                    g.Dispose();
                }

                thumbBitmap.Save(thumbPath, ImageFormat.Png);
            }
            finally
            {
                thumbBitmap.Dispose();
            }
        }
        finally
        {
            sourceBitmap.Dispose();
        }
    }
}
