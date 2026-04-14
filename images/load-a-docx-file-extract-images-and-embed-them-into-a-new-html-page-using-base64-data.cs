using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image (input.png) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(artifactsDir, "input.png");
        CreateSampleImage(imagePath);

        // -----------------------------------------------------------------
        // 2. Create a DOCX document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, imagePath);

        // -----------------------------------------------------------------
        // 3. Load the DOCX, extract all images and embed them into HTML.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "output.html");
        ExtractImagesToHtml(docPath, htmlPath);
    }

    private static void CreateSampleImage(string fileName)
    {
        // 200x200 white bitmap.
        Bitmap bitmap = new Bitmap(200, 200);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Simple deterministic drawing – a black rectangle.
        graphics.DrawRectangle(Pens.Black, 20, 20, 160, 160);
        // Save and clean up.
        bitmap.Save(fileName);
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateDocumentWithImage(string docFile, string imageFile)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image into the document.
        builder.InsertImage(imageFile);
        // Save the document.
        doc.Save(docFile);
    }

    private static void ExtractImagesToHtml(string docFile, string htmlFile)
    {
        Document doc = new Document(docFile);

        // Collect all shape nodes.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.AppendLine("<html>");
        htmlBuilder.AppendLine("<body>");

        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            ImageData imageData = shape.ImageData;

            using (MemoryStream ms = new MemoryStream())
            {
                // Save image bytes to the stream.
                imageData.Save(ms);
                // Reset position before reading.
                ms.Position = 0;

                // Convert to Base64.
                string base64 = Convert.ToBase64String(ms.ToArray());

                // Determine MIME type from ImageType.
                string mime = GetMimeType(imageData.ImageType);

                // Append <img> tag with embedded data.
                htmlBuilder.AppendLine(
                    $"<img src=\"data:{mime};base64,{base64}\" alt=\"Image{extractedCount}\" />");

                extractedCount++;
            }
        }

        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Save the HTML file.
        File.WriteAllText(htmlFile, htmlBuilder.ToString());
    }

    private static string GetMimeType(ImageType imageType)
    {
        // Map Aspose.Words.ImageType to standard MIME types.
        switch (imageType)
        {
            case ImageType.Jpeg:
                return "image/jpeg";
            case ImageType.Png:
                return "image/png";
            case ImageType.Bmp:
                return "image/bmp";
            case ImageType.Gif:
                return "image/gif";
            case ImageType.WebP:
                return "image/webp";
            case ImageType.Emf:
                return "image/emf";
            case ImageType.Wmf:
                return "image/wmf";
            case ImageType.Pict:
                return "image/pict";
            case ImageType.Eps:
                return "image/eps";
            default:
                return "application/octet-stream";
        }
    }
}
