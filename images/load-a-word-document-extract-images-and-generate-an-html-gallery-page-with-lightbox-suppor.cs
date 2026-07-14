using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base output directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string imagesDir = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(imagesDir);

        // 1. Create a deterministic sample image using Aspose.Drawing.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
        {
            g.DrawRectangle(pen, 10, 10, 180, 180);
        }
        bitmap.Save(sampleImagePath);
        g.Dispose();
        bitmap.Dispose();

        // 2. Build a Word document that contains the sample image multiple times.
        string docPath = Path.Combine(baseDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with images:");
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln();
        }
        doc.Save(docPath);

        // 3. Load the document.
        Document loadedDoc = new Document(docPath);

        // 4. Extract images from the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 5. Generate a simple HTML gallery with lightbox-like behavior.
        string htmlPath = Path.Combine(baseDir, "gallery.html");
        StringBuilder html = new StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html lang=\"en\">");
        html.AppendLine("<head>");
        html.AppendLine("    <meta charset=\"UTF-8\">");
        html.AppendLine("    <title>Image Gallery</title>");
        // Minimal CSS for a lightbox effect.
        html.AppendLine("    <style>");
        html.AppendLine("        .gallery { display: flex; flex-wrap: wrap; gap: 10px; }");
        html.AppendLine("        .gallery a { display: block; width: 150px; }");
        html.AppendLine("        .gallery img { width: 100%; height: auto; border: 1px solid #ccc; }");
        html.AppendLine("        .lightbox {");
        html.AppendLine("            position: fixed; top: 0; left: 0; width: 100%; height: 100%;");
        html.AppendLine("            background: rgba(0,0,0,0.8); display: none; align-items: center; justify-content: center;");
        html.AppendLine("        }");
        html.AppendLine("        .lightbox img { max-width: 90%; max-height: 90%; }");
        html.AppendLine("    </style>");
        html.AppendLine("    <script>");
        html.AppendLine("        function showLightbox(src) {");
        html.AppendLine("            var lb = document.getElementById('lightbox');");
        html.AppendLine("            var img = lb.querySelector('img');");
        html.AppendLine("            img.src = src;");
        html.AppendLine("            lb.style.display = 'flex';");
        html.AppendLine("        }");
        html.AppendLine("        function hideLightbox() { document.getElementById('lightbox').style.display = 'none'; }");
        html.AppendLine("    </script>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine("    <h1>Image Gallery</h1>");
        html.AppendLine("    <div class=\"gallery\">");

        // Add each extracted image to the gallery.
        for (int i = 0; i < imageIndex; i++)
        {
            string imgFile = $"Images/image_{i}{Path.GetExtension(Directory.GetFiles(imagesDir)[i])}";
            string imgPath = Path.Combine("Images", $"image_{i}{Path.GetExtension(Directory.GetFiles(imagesDir)[i])}");
            html.AppendLine($"        <a href=\"javascript:void(0);\" onclick=\"showLightbox('{imgPath}')\">");
            html.AppendLine($"            <img src=\"{imgPath}\" alt=\"Image {i}\" />");
            html.AppendLine("        </a>");
        }

        html.AppendLine("    </div>");
        html.AppendLine("    <div id=\"lightbox\" class=\"lightbox\" onclick=\"hideLightbox()\">");
        html.AppendLine("        <img src=\"\" alt=\"Lightbox\" />");
        html.AppendLine("    </div>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        File.WriteAllText(htmlPath, html.ToString(), Encoding.UTF8);
    }
}
