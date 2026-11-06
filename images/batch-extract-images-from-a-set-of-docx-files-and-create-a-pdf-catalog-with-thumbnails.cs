using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define deterministic folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string extractedFolder = Path.Combine(baseDir, "ExtractedImages");
        string catalogPath = Path.Combine(baseDir, "Catalog.pdf");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure clean environment
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (PNG)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle with a color
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the image
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            // Insert the sample image
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Extract images from each DOCX file
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document srcDoc = new Document(docFile);
            NodeCollection shapes = srcDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}.png";
                    string imagePath = Path.Combine(extractedFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // -------------------------------------------------
        // 4. Create a PDF catalog with thumbnails of extracted images
        // -------------------------------------------------
        Document catalogDoc = new Document();
        DocumentBuilder catalogBuilder = new DocumentBuilder(catalogDoc);

        // Title
        catalogBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        catalogBuilder.Font.Size = 24;
        catalogBuilder.Writeln("Image Catalog");
        catalogBuilder.Font.Size = 12;
        catalogBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        catalogBuilder.Writeln();

        // Insert each thumbnail
        foreach (string imgPath in extractedImagePaths)
        {
            // Caption
            catalogBuilder.Writeln(Path.GetFileName(imgPath));
            // Insert thumbnail (width 100 points, height auto)
            catalogBuilder.InsertImage(imgPath, 100, 0);
            catalogBuilder.Writeln(); // Add spacing
        }

        // Save as PDF
        catalogDoc.Save(catalogPath, SaveFormat.Pdf);

        // Validate output PDF exists
        if (!File.Exists(catalogPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");

        // -------------------------------------------------
        // 5. Clean up resources (optional, as using statements handle most)
        // -------------------------------------------------
        // All disposable objects have been disposed via using statements.
    }
}
