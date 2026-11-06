using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class ImageGalleryGenerator
{
    public static void Main()
    {
        // Define file and folder names.
        string sampleImagePath = "sample.png";
        string documentPath = "sample.docx";
        string imagesFolder = "extracted_images";
        string htmlPath = "gallery.html";

        // Ensure a clean environment.
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);
        if (File.Exists(documentPath)) File.Delete(documentPath);
        if (File.Exists(htmlPath)) File.Delete(htmlPath);
        if (Directory.Exists(imagesFolder)) Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        // -------------------------------------------------
        // Create a deterministic sample image using Aspose.Drawing.
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            // Save the image to a deterministic file.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Create a Word document and insert the sample image multiple times.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert three images.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath);
        }
        // Save the document.
        doc.Save(documentPath);

        // -------------------------------------------------
        // Load the document (simulating a separate load step).
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);

        // -------------------------------------------------
        // Extract images from the document.
        // -------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outImagePath = Path.Combine(imagesFolder, $"image{imageIndex}.png");
                shape.ImageData.Save(outImagePath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // Generate a simple HTML gallery with lightbox support.
        // -------------------------------------------------
        using (StreamWriter writer = new StreamWriter(htmlPath))
        {
            writer.WriteLine("<!DOCTYPE html>");
            writer.WriteLine("<html lang=\"en\">");
            writer.WriteLine("<head>");
            writer.WriteLine("    <meta charset=\"UTF-8\">");
            writer.WriteLine("    <title>Image Gallery</title>");
            writer.WriteLine("    <style>");
            writer.WriteLine("        body { font-family: Arial, sans-serif; margin: 0; padding: 0; }");
            writer.WriteLine("        .gallery { display: flex; flex-wrap: wrap; padding: 10px; }");
            writer.WriteLine("        .gallery img { margin: 5px; width: 150px; cursor: pointer; border: 2px solid #ccc; }");
            writer.WriteLine("        #lightboxOverlay {");
            writer.WriteLine("            position: fixed; top: 0; left: 0; width: 100%; height: 100%;");
            writer.WriteLine("            background: rgba(0,0,0,0.8); display: none; align-items: center; justify-content: center;");
            writer.WriteLine("        }");
            writer.WriteLine("        #lightboxOverlay img { max-width: 90%; max-height: 90%; }");
            writer.WriteLine("    </style>");
            writer.WriteLine("    <script>");
            writer.WriteLine("        function openLightbox(src) {");
            writer.WriteLine("            var overlay = document.getElementById('lightboxOverlay');");
            writer.WriteLine("            var img = overlay.querySelector('img');");
            writer.WriteLine("            img.src = src;");
            writer.WriteLine("            overlay.style.display = 'flex';");
            writer.WriteLine("        }");
            writer.WriteLine("        function closeLightbox() {");
            writer.WriteLine("            document.getElementById('lightboxOverlay').style.display = 'none';");
            writer.WriteLine("        }");
            writer.WriteLine("    </script>");
            writer.WriteLine("</head>");
            writer.WriteLine("<body>");
            writer.WriteLine("    <div class=\"gallery\">");

            // Add each extracted image to the gallery.
            for (int i = 0; i < imageIndex; i++)
            {
                string relPath = Path.Combine(imagesFolder, $"image{i}.png").Replace("\\", "/");
                writer.WriteLine($"        <img src=\"{relPath}\" onclick=\"openLightbox('{relPath}')\" alt=\"Image {i}\">");
            }

            writer.WriteLine("    </div>");
            writer.WriteLine("    <div id=\"lightboxOverlay\" onclick=\"closeLightbox()\">");
            writer.WriteLine("        <img src=\"\" alt=\"Lightbox Image\">");
            writer.WriteLine("    </div>");
            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
        }

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Failed to create the HTML gallery page.");
    }
}
