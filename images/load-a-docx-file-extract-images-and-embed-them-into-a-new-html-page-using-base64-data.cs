using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // File paths
        string imagePath = Path.Combine(artifactsDir, "input.png");
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        string htmlPath = Path.Combine(artifactsDir, "output.html");

        // 1. Create a deterministic sample image
        CreateSampleImage(imagePath);

        // 2. Create a DOCX and insert the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // 3. Load the DOCX
        Document loadedDoc = new Document(docPath);

        // 4. Extract images and embed them as Base64 in HTML
        StringBuilder html = new StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html><head><meta charset=\"UTF-8\"><title>Extracted Images</title></head><body>");

        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Get raw image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Determine MIME type from image extension
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType).ToLowerInvariant();
            string mime = GetMimeType(extension);

            // Convert to Base64 and write <img> tag
            string base64 = Convert.ToBase64String(imageBytes);
            html.AppendLine($"<img src=\"data:{mime};base64,{base64}\" alt=\"Image{imageCount}\"/>");
            imageCount++;
        }

        html.AppendLine("</body></html>");

        // Validate that at least one image was extracted
        if (imageCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 5. Save the HTML file
        File.WriteAllText(htmlPath, html.ToString());
    }

    // Creates a simple PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath)
    {
        int width = 200;
        int height = 100;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath);
        }
    }

    // Maps file extensions to MIME types
    private static string GetMimeType(string extension)
    {
        switch (extension)
        {
            case ".png":  return "image/png";
            case ".jpg":
            case ".jpeg": return "image/jpeg";
            case ".gif":  return "image/gif";
            case ".bmp":  return "image/bmp";
            case ".webp": return "image/webp";
            case ".tif":
            case ".tiff": return "image/tiff";
            default:      return "application/octet-stream";
        }
    }
}
