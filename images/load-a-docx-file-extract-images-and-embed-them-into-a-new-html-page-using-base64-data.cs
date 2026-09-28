using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string imagePath = "input.png";
        const string docPath = "sample.docx";
        const string htmlPath = "output.html";

        // -------------------------------------------------
        // 1. Create a deterministic sample image (100x100 white)
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle for visual distinction
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(imagePath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and insert the sample image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the DOCX document and extract images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var htmlBuilder = new StringBuilder();

        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html>");
        htmlBuilder.AppendLine("<head><meta charset=\"UTF-8\"><title>Extracted Images</title></head>");
        htmlBuilder.AppendLine("<body>");

        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Get raw image bytes
            byte[] imageBytes = shape.ImageData.ImageBytes;

            // Determine MIME type based on image format
            string mimeType = "application/octet-stream";
            switch (shape.ImageData.ImageType)
            {
                case ImageType.Jpeg:
                    mimeType = "image/jpeg";
                    break;
                case ImageType.Png:
                    mimeType = "image/png";
                    break;
                case ImageType.Gif:
                    mimeType = "image/gif";
                    break;
                case ImageType.Bmp:
                    mimeType = "image/bmp";
                    break;
                case ImageType.Emf:
                    mimeType = "image/emf";
                    break;
                case ImageType.Wmf:
                    mimeType = "image/wmf";
                    break;
                // Additional types can be added here if needed
            }

            // Convert to Base64
            string base64 = Convert.ToBase64String(imageBytes);

            // Embed into HTML
            htmlBuilder.AppendLine($"<img src=\"data:{mimeType};base64,{base64}\" alt=\"Extracted Image\" />");
            extractedCount++;
        }

        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Write the HTML content to a file
        // -------------------------------------------------
        File.WriteAllText(htmlPath, htmlBuilder.ToString(), Encoding.UTF8);

        // Validate that the HTML file was created
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Failed to create the output HTML file.");

        // Optional cleanup (commented out for inspection)
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
