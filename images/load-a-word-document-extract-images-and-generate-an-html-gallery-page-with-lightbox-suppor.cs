using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images using Aspose.Drawing.
        // -----------------------------------------------------------------
        string image1Path = Path.Combine(baseDir, "sample1.png");
        string image2Path = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(image1Path, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(image2Path, 150, 200, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains the sample images.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(baseDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertImage(image1Path);
        builder.Writeln(); // separate the images with a line break
        builder.InsertImage(image2Path);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all embedded images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        string imagesFolder = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesFolder);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image{extractedCount}{extension}";
                string imageFullPath = Path.Combine(imagesFolder, imageFileName);
                shape.ImageData.Save(imageFullPath);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Generate a simple HTML gallery with lightbox support.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(baseDir, "gallery.html");
        string htmlContent = BuildHtmlGallery(imagesFolder, extractedCount);
        File.WriteAllText(htmlPath, htmlContent, Encoding.UTF8);

        // Execution finished – all files are written to the Artifacts folder.
    }

    // Creates a PNG image with a solid background colour.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath);
        }
    }

    // Builds the HTML string for the gallery page.
    private static string BuildHtmlGallery(string imagesFolder, int imageCount)
    {
        // Relative path from the HTML file to the images folder.
        string relativeImagesPath = "Images";

        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset=\"UTF-8\">");
        sb.AppendLine("    <title>Image Gallery</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: Arial, sans-serif; background:#f0f0f0; margin:0; padding:20px; }");
        sb.AppendLine("        .gallery { display:flex; flex-wrap:wrap; gap:10px; }");
        sb.AppendLine("        .thumb { width:150px; height:auto; cursor:pointer; border:2px solid #fff; box-shadow:0 2px 5px rgba(0,0,0,0.3); }");
        sb.AppendLine("        .lightbox { display:none; position:fixed; top:0; left:0; width:100%; height:100%;");
        sb.AppendLine("                    background:rgba(0,0,0,0.8); align-items:center; justify-content:center; }");
        sb.AppendLine("        .lightbox img { max-width:90%; max-height:90%; }");
        sb.AppendLine("        .lightbox:target { display:flex; }");
        sb.AppendLine("    </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("    <h1>Image Gallery</h1>");
        sb.AppendLine("    <div class=\"gallery\">");

        for (int i = 0; i < imageCount; i++)
        {
            // Determine the file name with unknown extension – we will match any file that starts with image{i}
            string pattern = $"image{i}";
            string[] matchingFiles = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Artifacts", "Images"))
                                             .Where(f => Path.GetFileName(f).StartsWith(pattern, StringComparison.OrdinalIgnoreCase))
                                             .ToArray();

            if (matchingFiles.Length == 0)
                continue; // safety check

            string fileName = Path.GetFileName(matchingFiles[0]);
            string thumbPath = $"{relativeImagesPath}/{fileName}";
            string lightboxId = $"lightbox{i}";

            sb.AppendLine($"        <a href=\"#{lightboxId}\"><img src=\"{thumbPath}\" class=\"thumb\" alt=\"Image {i}\"/></a>");
            sb.AppendLine($"        <div id=\"{lightboxId}\" class=\"lightbox\">");
            sb.AppendLine($"            <a href=\"#\" style=\"position:absolute;top:20px;right:30px;color:#fff;font-size:30px;text-decoration:none;\">&times;</a>");
            sb.AppendLine($"            <img src=\"{thumbPath}\" alt=\"Image {i}\"/>");
            sb.AppendLine("        </div>");
        }

        sb.AppendLine("    </div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }
}
