using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create sample images using Aspose.Drawing.
        // -----------------------------------------------------------------
        string[] sampleImageFiles = { Path.Combine(outputDir, "sample1.png"), Path.Combine(outputDir, "sample2.png") };
        CreateSampleImage(sampleImageFiles[0], 200, 200, Aspose.Drawing.Color.LightBlue, "A");
        CreateSampleImage(sampleImageFiles[1], 200, 200, Aspose.Drawing.Color.LightCoral, "B");

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample images.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in sampleImageFiles)
        {
            builder.InsertImage(imgPath);
            builder.Writeln(); // separate images with a line break.
        }

        doc.Save(docPath); // Save the document.

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        string imagesDir = Path.Combine(outputDir, "extracted");
        Directory.CreateDirectory(imagesDir);

        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string fullPath = Path.Combine(imagesDir, imageFileName);
                shape.ImageData.Save(fullPath);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validation: at least one image must have been extracted.
        if (extractedImageFiles.Count == 0)
            throw new Exception("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Generate a simple HTML gallery with Lightbox support.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(outputDir, "gallery.html");
        string htmlContent = GenerateHtmlGallery(extractedImageFiles, "extracted");
        File.WriteAllText(htmlPath, htmlContent);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color background, string label)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(background);
        // Simple text drawing to differentiate images.
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
        {
            graphics.DrawString(label, font, Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(50, 80));
        }
        graphics.Dispose();
        bitmap.Save(filePath);
        bitmap.Dispose();
    }

    // Builds the HTML string for the gallery.
    private static string GenerateHtmlGallery(List<string> imageFiles, string imagesFolder)
    {
        // Lightbox2 CDN links (no runtime download, just references).
        const string cssCdn = "https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/css/lightbox.min.css";
        const string jsCdn = "https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/js/lightbox.min.js";

        var html = new System.Text.StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html lang=\"en\">");
        html.AppendLine("<head>");
        html.AppendLine("    <meta charset=\"UTF-8\">");
        html.AppendLine("    <title>Image Gallery</title>");
        html.AppendLine($"    <link rel=\"stylesheet\" href=\"{cssCdn}\">");
        html.AppendLine("    <style>");
        html.AppendLine("        .gallery { display: flex; flex-wrap: wrap; gap: 10px; }");
        html.AppendLine("        .gallery a { width: 150px; }");
        html.AppendLine("        .gallery img { width: 100%; height: auto; border: 1px solid #ccc; }");
        html.AppendLine("    </style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine("    <h1>Image Gallery</h1>");
        html.AppendLine("    <div class=\"gallery\">");

        foreach (string fileName in imageFiles)
        {
            string relativePath = $"{imagesFolder}/{fileName}";
            html.AppendLine($"        <a href=\"{relativePath}\" data-lightbox=\"gallery\"><img src=\"{relativePath}\" alt=\"{fileName}\"/></a>");
        }

        html.AppendLine("    </div>");
        html.AppendLine($"    <script src=\"{jsCdn}\"></script>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        return html.ToString();
    }
}
