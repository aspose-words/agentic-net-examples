using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace ImageGalleryExample
{
    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a deterministic sample image (input.png)
            const string sampleImagePath = "input.png";
            CreateSampleImage(sampleImagePath);

            // Step 2: Create a Word document and insert the sample image multiple times
            const string docPath = "sample.docx";
            CreateWordDocumentWithImages(docPath, sampleImagePath);

            // Step 3: Load the Word document
            Document doc = new Document(docPath);

            // Step 4: Extract images from the document
            List<string> extractedImages = ExtractImagesFromDocument(doc);

            // Validate that at least one image was extracted
            if (extractedImages.Count == 0)
                throw new InvalidOperationException("No images were extracted from the document.");

            // Step 5: Generate an HTML gallery page with Lightbox support
            const string htmlPath = "gallery.html";
            GenerateHtmlGallery(htmlPath, extractedImages);

            // Validate that the HTML file was created
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException("Failed to create the HTML gallery page.");
        }

        private static void CreateSampleImage(string path)
        {
            // Create a 200x200 white bitmap and draw a simple rectangle
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                    // Draw a blue rectangle
                    using (var pen = new Aspose.Drawing.Pen(Color.Blue, 5))
                    {
                        graphics.DrawRectangle(pen, 25, 25, 150, 150);
                    }
                }
                bitmap.Save(path);
            }

            // Ensure the image file exists
            if (!File.Exists(path))
                throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
        }

        private static void CreateWordDocumentWithImages(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the image three times
            for (int i = 0; i < 3; i++)
            {
                builder.InsertParagraph();
                builder.InsertImage(imagePath);
            }

            doc.Save(docPath);

            // Ensure the document file exists
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create Word document at '{docPath}'.");
        }

        private static List<string> ExtractImagesFromDocument(Document doc)
        {
            List<string> imagePaths = new List<string>();
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int index = 1;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"image{index}.png";
                    shape.ImageData.Save(imageFileName);
                    imagePaths.Add(imageFileName);
                    index++;
                }
            }

            return imagePaths;
        }

        private static void GenerateHtmlGallery(string htmlPath, List<string> imageFiles)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"en\">");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset=\"UTF-8\">");
            sb.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine("    <title>Image Gallery</title>");
            sb.AppendLine("    <link href=\"https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/css/lightbox.min.css\" rel=\"stylesheet\" />");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("    <h1>Extracted Images Gallery</h1>");
            sb.AppendLine("    <div>");

            foreach (string img in imageFiles)
            {
                // Use the same image for thumbnail; style limits size
                sb.AppendLine($"        <a href=\"{img}\" data-lightbox=\"gallery\"><img src=\"{img}\" style=\"width:200px;height:auto;margin:5px;\" /></a>");
            }

            sb.AppendLine("    </div>");
            sb.AppendLine("    <script src=\"https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/js/lightbox-plus-jquery.min.js\"></script>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            File.WriteAllText(htmlPath, sb.ToString());

            // Ensure the HTML file exists
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to write HTML gallery to '{htmlPath}'.");
        }
    }
}
