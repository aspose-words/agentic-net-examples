using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Helper: creates a deterministic PNG image.
    private static void CreateSamplePng(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        // Create bitmap.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height, Aspose.Drawing.Imaging.PixelFormat.Format32bppArgb))
        {
            // Fill background.
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);

                // Draw centered text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
                {
                    // Measure text size.
                    SizeF textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2f;
                    float y = (height - textSize.Height) / 2f;
                    graphics.DrawString(text, font, Aspose.Drawing.Brushes.Black, x, y);
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    public static void Main()
    {
        // Define folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string imagesInputDir = Path.Combine(artifactsDir, "InputImages");
        string imagesExtractDir = Path.Combine(artifactsDir, "ExtractedImages");
        string htmlOutputPath = Path.Combine(artifactsDir, "gallery.html");
        string docPath = Path.Combine(artifactsDir, "sample.docx");

        // Ensure clean environment.
        if (Directory.Exists(artifactsDir))
            Directory.Delete(artifactsDir, true);
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesInputDir);
        Directory.CreateDirectory(imagesExtractDir);

        // -------------------------------------------------
        // 1. Create deterministic sample images (PNG).
        // -------------------------------------------------
        string sampleImagePath1 = Path.Combine(imagesInputDir, "sample1.png");
        CreateSamplePng(sampleImagePath1, 200, 200, Aspose.Drawing.Color.LightBlue, "Img1");

        string sampleImagePath2 = Path.Combine(imagesInputDir, "sample2.png");
        CreateSamplePng(sampleImagePath2, 200, 200, Aspose.Drawing.Color.LightCoral, "Img2");

        // -------------------------------------------------
        // 2. Create a Word document and insert the images.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Image Gallery Example");
        builder.InsertParagraph();

        // Insert first image.
        builder.InsertImage(sampleImagePath1);
        builder.InsertParagraph();

        // Insert second image.
        builder.InsertImage(sampleImagePath2);
        builder.InsertParagraph();

        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        List<string> extractedImageFiles = new List<string>();

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedFileName = $"image_{imageIndex}{extension}";
                string extractedFullPath = Path.Combine(imagesExtractDir, extractedFileName);

                shape.ImageData.Save(extractedFullPath);
                extractedImageFiles.Add(extractedFileName);
                imageIndex++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Generate HTML gallery with simple lightbox effect.
        // -------------------------------------------------
        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html lang=\"en\">");
        htmlBuilder.AppendLine("<head>");
        htmlBuilder.AppendLine("    <meta charset=\"UTF-8\">");
        htmlBuilder.AppendLine("    <title>Image Gallery</title>");
        htmlBuilder.AppendLine("    <style>");
        htmlBuilder.AppendLine("        .gallery { display: flex; flex-wrap: wrap; gap: 10px; }");
        htmlBuilder.AppendLine("        .gallery a { border: 1px solid #ccc; }");
        htmlBuilder.AppendLine("        .gallery img { display: block; max-width: 200px; height: auto; }");
        htmlBuilder.AppendLine("        .lightbox {");
        htmlBuilder.AppendLine("            position: fixed; top: 0; left: 0; width: 100%; height: 100%;");
        htmlBuilder.AppendLine("            background: rgba(0,0,0,0.8); display: none; align-items: center; justify-content: center;");
        htmlBuilder.AppendLine("        }");
        htmlBuilder.AppendLine("        .lightbox img { max-width: 90%; max-height: 90%; }");
        htmlBuilder.AppendLine("    </style>");
        htmlBuilder.AppendLine("</head>");
        htmlBuilder.AppendLine("<body>");
        htmlBuilder.AppendLine("    <h1>Extracted Images Gallery</h1>");
        htmlBuilder.AppendLine("    <div class=\"gallery\">");

        foreach (string fileName in extractedImageFiles)
        {
            string relativePath = Path.Combine("ExtractedImages", fileName).Replace("\\", "/");
            htmlBuilder.AppendLine($"        <a href=\"{relativePath}\" onclick=\"showLightbox(event, '{relativePath}'); return false;\">");
            htmlBuilder.AppendLine($"            <img src=\"{relativePath}\" alt=\"Image\" />");
            htmlBuilder.AppendLine("        </a>");
        }

        htmlBuilder.AppendLine("    </div>");
        htmlBuilder.AppendLine("    <div class=\"lightbox\" id=\"lightbox\" onclick=\"hideLightbox()\">");
        htmlBuilder.AppendLine("        <img id=\"lightboxImg\" src=\"\" alt=\"\" />");
        htmlBuilder.AppendLine("    </div>");
        htmlBuilder.AppendLine("    <script>");
        htmlBuilder.AppendLine("        function showLightbox(event, src) {");
        htmlBuilder.AppendLine("            var lightbox = document.getElementById('lightbox');");
        htmlBuilder.AppendLine("            var img = document.getElementById('lightboxImg');");
        htmlBuilder.AppendLine("            img.src = src;");
        htmlBuilder.AppendLine("            lightbox.style.display = 'flex';");
        htmlBuilder.AppendLine("        }");
        htmlBuilder.AppendLine("        function hideLightbox() {");
        htmlBuilder.AppendLine("            var lightbox = document.getElementById('lightbox');");
        htmlBuilder.AppendLine("            lightbox.style.display = 'none';");
        htmlBuilder.AppendLine("        }");
        htmlBuilder.AppendLine("    </script>");
        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");

        // Write HTML file.
        File.WriteAllText(htmlOutputPath, htmlBuilder.ToString(), Encoding.UTF8);

        // Final validation: ensure HTML file exists.
        if (!File.Exists(htmlOutputPath))
            throw new InvalidOperationException("Failed to create the HTML gallery page.");
    }
}
