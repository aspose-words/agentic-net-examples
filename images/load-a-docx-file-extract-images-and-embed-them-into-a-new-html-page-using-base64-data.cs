using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Use explicit Aspose.Drawing types as required by the rules.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        try
        {
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            finally
            {
                graphics.Dispose();
            }

            bitmap.Save(sampleImagePath);
        }
        finally
        {
            bitmap.Dispose();
        }

        // Create a DOCX document and insert the sample image.
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // Load the DOCX file.
        Document loadedDoc = new Document(inputDocPath);

        // Extract images and convert them to base64 data URIs.
        List<string> imageDataUris = new List<string>();
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;
                    string base64 = Convert.ToBase64String(ms.ToArray());
                    string mime = GetMimeType(shape.ImageData.ImageType);
                    imageDataUris.Add($"data:{mime};base64,{base64}");
                }
            }
        }

        // Validate that at least one image was extracted.
        if (imageDataUris.Count == 0)
            throw new Exception("No images were extracted from the document.");

        // Build an HTML page embedding the images.
        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html>");
        htmlBuilder.AppendLine("<head>");
        htmlBuilder.AppendLine("<meta charset=\"UTF-8\">");
        htmlBuilder.AppendLine("<title>Extracted Images</title>");
        htmlBuilder.AppendLine("</head>");
        htmlBuilder.AppendLine("<body>");
        foreach (string dataUri in imageDataUris)
        {
            htmlBuilder.AppendLine($"<img src=\"{dataUri}\" alt=\"Extracted Image\"/>");
        }
        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");

        // Save the HTML page.
        const string outputHtmlPath = "output.html";
        File.WriteAllText(outputHtmlPath, htmlBuilder.ToString());

        // Validate that the HTML file was created.
        if (!File.Exists(outputHtmlPath))
            throw new Exception("Failed to create the output HTML file.");
    }

    // Helper method to map Aspose.Words.Drawing.ImageType to MIME types.
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
}
