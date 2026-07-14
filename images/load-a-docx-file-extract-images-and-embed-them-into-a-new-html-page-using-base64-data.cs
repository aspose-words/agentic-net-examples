using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string imagePath = "sample.png";
        const string docPath = "input.docx";
        const string htmlPath = "output.html";

        // -------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            // Save the image to a local file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the DOCX document.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -------------------------------------------------
        // 4. Extract images and embed them as Base64 in HTML.
        // -------------------------------------------------
        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html>");
        htmlBuilder.AppendLine("<head><meta charset=\"UTF-8\"><title>Extracted Images</title></head>");
        htmlBuilder.AppendLine("<body>");

        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get raw image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            if (imageBytes == null || imageBytes.Length == 0)
                continue;

            // Determine MIME type from image format.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType); // e.g., ".png"
            string mimeType = "image/" + extension.TrimStart('.').ToLowerInvariant();

            // Encode to Base64.
            string base64 = Convert.ToBase64String(imageBytes);

            // Append <img> tag.
            htmlBuilder.AppendLine($"<img src=\"data:{mimeType};base64,{base64}\" alt=\"Extracted Image\" />");
            extractedCount++;
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");

        // Save the HTML file.
        File.WriteAllText(htmlPath, htmlBuilder.ToString(), Encoding.UTF8);
    }
}
