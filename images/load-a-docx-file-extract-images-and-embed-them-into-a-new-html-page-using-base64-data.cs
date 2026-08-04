using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Step 1: Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Create bitmap and fill with a solid color.
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Dispose drawing objects.
        graphics.Dispose();
        bitmap.Save(sampleImagePath);
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // Step 2: Create a DOCX document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Step 3: Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // Step 4: Extract images from the document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> base64Images = new List<string>();
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get raw image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            if (imageBytes == null || imageBytes.Length == 0)
                continue;

            // Determine MIME type from image extension.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType).ToLowerInvariant(); // e.g., ".png"
            string mime;
            switch (extension)
            {
                case ".png":
                    mime = "image/png";
                    break;
                case ".jpeg":
                case ".jpg":
                    mime = "image/jpeg";
                    break;
                case ".gif":
                    mime = "image/gif";
                    break;
                case ".bmp":
                    mime = "image/bmp";
                    break;
                case ".webp":
                    mime = "image/webp";
                    break;
                case ".tiff":
                case ".tif":
                    mime = "image/tiff";
                    break;
                default:
                    mime = "application/octet-stream";
                    break;
            }

            // Convert to Base64 and build data URI.
            string base64 = Convert.ToBase64String(imageBytes);
            string dataUri = $"data:{mime};base64,{base64}";
            base64Images.Add(dataUri);
        }

        // Validate that at least one image was extracted.
        if (base64Images.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // Step 5: Generate HTML with embedded Base64 images.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "output.html");
        using (StreamWriter writer = new StreamWriter(htmlPath, false))
        {
            writer.WriteLine("<!DOCTYPE html>");
            writer.WriteLine("<html>");
            writer.WriteLine("<head><meta charset=\"UTF-8\"><title>Extracted Images</title></head>");
            writer.WriteLine("<body>");
            foreach (string dataUri in base64Images)
            {
                writer.WriteLine($"<img src=\"{dataUri}\" alt=\"Embedded Image\" style=\"margin:10px;\" />");
            }
            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
        }

        // -----------------------------------------------------------------
        // Completion message (optional, not required for non‑interactive run).
        // -----------------------------------------------------------------
        Console.WriteLine($"HTML file with embedded images created at: {htmlPath}");
    }
}
