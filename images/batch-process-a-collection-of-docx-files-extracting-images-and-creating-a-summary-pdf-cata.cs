using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders and files
        string baseDir = Directory.GetCurrentDirectory();
        string docsFolder = Path.Combine(baseDir, "Docs");
        string imagesFolder = Path.Combine(baseDir, "ExtractedImages");
        string catalogPdfPath = Path.Combine(baseDir, "ImageCatalog.pdf");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure clean environment
        if (Directory.Exists(docsFolder)) Directory.Delete(docsFolder, true);
        if (Directory.Exists(imagesFolder)) Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(docsFolder);
        Directory.CreateDirectory(imagesFolder);

        // -------------------------------------------------
        // Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Create sample DOCX files each containing the image
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} with an embedded image:");
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(docsFolder, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // Batch process: extract images from each DOCX
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        string[] docFiles = Directory.GetFiles(docsFolder, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    // Use a deterministic PNG extension for extracted images
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}.png";
                    string imagePath = Path.Combine(imagesFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that images were extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // -------------------------------------------------
        // Create a summary PDF catalog containing all extracted images
        // -------------------------------------------------
        Document catalogDoc = new Document();
        DocumentBuilder catalogBuilder = new DocumentBuilder(catalogDoc);
        catalogBuilder.Writeln("Image Catalog");
        catalogBuilder.Font.Size = 24;
        catalogBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        catalogBuilder.Writeln();
        catalogBuilder.Font.Size = 12;
        catalogBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        bool first = true;
        foreach (string imgPath in extractedImagePaths)
        {
            if (!first)
                catalogBuilder.InsertBreak(BreakType.PageBreak);
            first = false;

            // Caption
            catalogBuilder.Writeln($"Image: {Path.GetFileName(imgPath)}");
            // Insert the image
            catalogBuilder.InsertImage(imgPath);
        }

        // Save the catalog as PDF
        catalogDoc.Save(catalogPdfPath, SaveFormat.Pdf);

        // Validate PDF creation
        if (!File.Exists(catalogPdfPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");

        // Optional: Output a simple JSON manifest
        var manifest = new
        {
            GeneratedAt = DateTime.UtcNow,
            SourceDocuments = docFiles,
            ExtractedImages = extractedImagePaths,
            CatalogPdf = catalogPdfPath
        };
        string manifestJson = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(Path.Combine(baseDir, "manifest.json"), manifestJson);
    }
}
