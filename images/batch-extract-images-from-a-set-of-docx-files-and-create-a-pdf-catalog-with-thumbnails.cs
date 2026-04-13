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
        // Prepare directories
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string docsDir = Path.Combine(baseDir, "InputDocs");
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);

        // Create a sample image to be used in documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files with the image inserted
        int docCount = 2;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            // Insert the sample image
            Shape imgShape = builder.InsertImage(sampleImagePath);
            imgShape.Width = 150;
            imgShape.Height = 150;
            string docPath = Path.Combine(docsDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // Extract images from all DOCX files
        List<string> extractedImagePaths = new List<string>();
        string[] docFiles = Directory.GetFiles(docsDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imgIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imgFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imgIndex}.png";
                    string imgPath = Path.Combine(imagesDir, imgFileName);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        shape.ImageData.Save(ms);
                        ms.Position = 0;
                        using (FileStream fs = new FileStream(imgPath, FileMode.Create, FileAccess.Write))
                        {
                            ms.CopyTo(fs);
                        }
                    }
                    extractedImagePaths.Add(imgPath);
                    imgIndex++;
                }
            }
        }

        // Validate extraction
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Create PDF catalog with thumbnails
        Document catalog = new Document();
        DocumentBuilder catBuilder = new DocumentBuilder(catalog);
        foreach (string imgPath in extractedImagePaths)
        {
            catBuilder.Writeln(Path.GetFileName(imgPath));
            Shape thumb = catBuilder.InsertImage(imgPath);
            thumb.Width = 100;
            thumb.Height = 100;
            catBuilder.Writeln(); // Add spacing
        }

        string catalogPath = Path.Combine(baseDir, "Catalog.pdf");
        catalog.Save(catalogPath, SaveFormat.Pdf);

        // Validate catalog creation
        if (!File.Exists(catalogPath))
            throw new InvalidOperationException("Catalog PDF was not created.");
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);

        // Draw a simple rectangle
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
        {
            g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Draw some text
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
        {
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
            {
                g.DrawString("Sample", font, brush, new Aspose.Drawing.PointF(20, height / 2 - 12));
            }
        }

        bitmap.Save(path, ImageFormat.Png);

        // Cleanup
        g.Dispose();
        bitmap.Dispose();
    }
}
