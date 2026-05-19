using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageGalleryGenerator
{
    public static void Main()
    {
        // Define directories
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 1. Create sample images using Aspose.Drawing (deterministic local files)
        // -----------------------------------------------------------------
        string[] sampleImagePaths = new string[2];
        for (int i = 0; i < sampleImagePaths.Length; i++)
        {
            string imgPath = Path.Combine(artifactsDir, $"sample{i + 1}.png");
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Fill background with a distinct color
                    g.Clear(i == 0 ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightCoral);
                }
                // Save the bitmap as PNG
                bitmap.Save(imgPath, ImageFormat.Png);
            }
            sampleImagePaths[i] = imgPath;
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample images
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in sampleImagePaths)
        {
            // Insert image inline
            builder.InsertImage(imgPath);
            builder.Writeln(); // Add a line break between images
        }

        // Save the document
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image data to the file system
                shape.ImageData.Save(imageFullPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Generate an HTML gallery page with simple lightbox support
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "gallery.html");
        StringBuilder htmlBuilder = new StringBuilder();

        // Basic HTML structure with a minimal lightbox implementation using CSS
        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html lang=\"en\">");
        htmlBuilder.AppendLine("<head>");
        htmlBuilder.AppendLine("    <meta charset=\"UTF-8\">");
        htmlBuilder.AppendLine("    <title>Image Gallery</title>");
        htmlBuilder.AppendLine("    <style>");
        htmlBuilder.AppendLine("        body { font-family: Arial, sans-serif; margin: 0; padding: 0; }");
        htmlBuilder.AppendLine("        .gallery { display: flex; flex-wrap: wrap; gap: 10px; padding: 10px; }");
        htmlBuilder.AppendLine("        .gallery img { width: 150px; height: auto; cursor: pointer; border: 2px solid #ccc; }");
        htmlBuilder.AppendLine("        .lightbox {");
        htmlBuilder.AppendLine("            position: fixed; top: 0; left: 0; width: 100%; height: 100%;");
        htmlBuilder.AppendLine("            background: rgba(0,0,0,0.8); display: none; align-items: center; justify-content: center;");
        htmlBuilder.AppendLine("        }");
        htmlBuilder.AppendLine("        .lightbox img { max-width: 90%; max-height: 90%; }");
        htmlBuilder.AppendLine("        .lightbox:target { display: flex; }");
        htmlBuilder.AppendLine("    </style>");
        htmlBuilder.AppendLine("</head>");
        htmlBuilder.AppendLine("<body>");
        htmlBuilder.AppendLine("    <div class=\"gallery\">");

        // Add each extracted image to the gallery
        for (int i = 0; i < imageIndex; i++)
        {
            string ext = Path.GetExtension(Directory.GetFiles(imagesDir, $"image_{i}.*")[0]); // get actual extension
            string imgFile = $"Images/image_{i}{ext}";
            string lightboxId = $"lightbox_{i}";
            htmlBuilder.AppendLine($"        <a href=\"#{lightboxId}\"><img src=\"{imgFile}\" alt=\"Image {i}\" /></a>");
            htmlBuilder.AppendLine($"        <div id=\"{lightboxId}\" class=\"lightbox\">");
            htmlBuilder.AppendLine($"            <a href=\"#\" style=\"position:absolute;top:20px;right:30px;color:#fff;font-size:30px;text-decoration:none;\">&times;</a>");
            htmlBuilder.AppendLine($"            <img src=\"{imgFile}\" alt=\"Image {i}\" />");
            htmlBuilder.AppendLine("        </div>");
        }

        htmlBuilder.AppendLine("    </div>");
        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");

        // Write HTML file
        File.WriteAllText(htmlPath, htmlBuilder.ToString(), Encoding.UTF8);

        // -----------------------------------------------------------------
        // Execution finished – the gallery.html file can be opened in a browser.
        // -----------------------------------------------------------------
    }
}
