using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image file.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Step 2: Create a DOCX document and insert the sample image.
        const string inputDocPath = "sample.docx";
        CreateDocumentWithImage(inputDocPath, sampleImagePath);

        // Step 3: Load the DOCX document.
        Document doc = new Document(inputDocPath);

        // Step 4: Extract images and convert them to Base64 strings.
        List<(string MimeType, string Base64)> extractedImages = ExtractImagesAsBase64(doc);

        // Validate that at least one image was extracted.
        if (extractedImages.Count == 0)
            throw new Exception("No images were extracted from the document.");

        // Step 5: Build HTML content embedding the images.
        string htmlContent = BuildHtmlWithImages(extractedImages);

        // Step 6: Save the HTML to a file.
        const string outputHtmlPath = "output.html";
        File.WriteAllText(outputHtmlPath, htmlContent);

        // Validate that the HTML file was created.
        if (!File.Exists(outputHtmlPath))
            throw new Exception("Failed to create the output HTML file.");
    }

    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 200;

        // Create a bitmap using Aspose.Drawing.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            // Obtain a graphics object for drawing.
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            finally
            {
                graphics.Dispose();
            }

            // Save the bitmap to the specified path.
            bitmap.Save(path);
        }
        finally
        {
            bitmap.Dispose();
        }

        // Ensure the image file exists.
        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Ensure the image file exists before insertion.
        if (!File.Exists(imagePath))
            throw new Exception($"Image file '{imagePath}' does not exist.");

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Ensure the document file exists.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document at '{docPath}'.");
    }

    private static List<(string MimeType, string Base64)> ExtractImagesAsBase64(Document doc)
    {
        List<(string MimeType, string Base64)> result = new List<(string, string)>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;
                byte[] imageBytes = ms.ToArray();
                string base64 = Convert.ToBase64String(imageBytes);
                string mime = GetMimeType(shape.ImageData.ImageType);
                result.Add((mime, base64));
            }
        }
        return result;
    }

    private static string GetMimeType(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return "image/jpeg";
            case ImageType.Png:
                return "image/png";
            case ImageType.Gif:
                return "image/gif";
            case ImageType.Bmp:
                return "image/bmp";
            case ImageType.Emf:
                return "image/emf";
            case ImageType.Wmf:
                return "image/wmf";
            default:
                return "application/octet-stream";
        }
    }

    private static string BuildHtmlWithImages(List<(string MimeType, string Base64)> images)
    {
        using (StringWriter writer = new StringWriter())
        {
            writer.WriteLine("<!DOCTYPE html>");
            writer.WriteLine("<html>");
            writer.WriteLine("<head>");
            writer.WriteLine("<meta charset=\"UTF-8\">");
            writer.WriteLine("<title>Extracted Images</title>");
            writer.WriteLine("</head>");
            writer.WriteLine("<body>");
            foreach (var (mime, base64) in images)
            {
                writer.WriteLine($"<img src=\"data:{mime};base64,{base64}\" alt=\"Extracted Image\" />");
                writer.WriteLine("<br/>");
            }
            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
            return writer.ToString();
        }
    }
}
