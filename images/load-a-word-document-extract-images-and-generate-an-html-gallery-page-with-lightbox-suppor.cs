using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace AsposeWordsImageGallery
{
    public class Program
    {
        public static void Main()
        {
            // Paths for temporary files
            const string sampleImagePath = "input.png";
            const string docPath = "sample.docx";
            const string htmlPath = "gallery.html";

            // 1. Create a deterministic sample image (200x200 white background with a blue rectangle)
            CreateSampleImage(sampleImagePath);

            // 2. Create a Word document and insert the sample image multiple times
            CreateWordDocumentWithImages(docPath, sampleImagePath, insertCount: 3);

            // 3. Load the document and extract all embedded images
            List<string> extractedImagePaths = ExtractImagesFromDocument(docPath);

            // Validate that at least one image was extracted
            if (extractedImagePaths.Count == 0)
                throw new InvalidOperationException("No images were extracted from the document.");

            // 4. Generate an HTML gallery page with Lightbox support
            GenerateHtmlGallery(htmlPath, extractedImagePaths);

            // Validate that the HTML file was created
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML gallery at '{htmlPath}'.");

            // Program completed successfully
        }

        private static void CreateSampleImage(string filePath)
        {
            const int width = 200;
            const int height = 200;

            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill background with white
                    graphics.Clear(Aspose.Drawing.Color.White);

                    // Draw a blue rectangle
                    using (var brush = new SolidBrush(Aspose.Drawing.Color.Blue))
                    {
                        graphics.FillRectangle(brush, 20, 20, width - 40, height - 40);
                    }
                }

                // Save the bitmap to a PNG file
                bitmap.Save(filePath);
            }

            // Ensure the file exists
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
        }

        private static void CreateWordDocumentWithImages(string docPath, string imagePath, int insertCount)
        {
            // Ensure the source image exists
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < insertCount; i++)
            {
                // Insert the image
                builder.InsertImage(imagePath);
                // Add a line break after each image for readability
                builder.Writeln();
            }

            // Save the document
            doc.Save(docPath);

            // Validate that the document was saved
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create Word document at '{docPath}'.");
        }

        private static List<string> ExtractImagesFromDocument(string docPath)
        {
            if (!File.Exists(docPath))
                throw new FileNotFoundException($"Document file not found: {docPath}");

            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            List<string> extractedPaths = new List<string>();
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"extracted_{imageIndex}.png";
                    shape.ImageData.Save(imageFileName);
                    extractedPaths.Add(imageFileName);
                    imageIndex++;
                }
            }

            // Validate that each extracted file exists
            foreach (string path in extractedPaths)
            {
                if (!File.Exists(path))
                    throw new InvalidOperationException($"Failed to save extracted image at '{path}'.");
            }

            return extractedPaths;
        }

        private static void GenerateHtmlGallery(string htmlPath, List<string> imagePaths)
        {
            // Lightbox2 CDN links
            const string lightboxCss = "https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/css/lightbox.min.css";
            const string lightboxJs = "https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/js/lightbox.min.js";

            using (StreamWriter writer = new StreamWriter(htmlPath, false))
            {
                writer.WriteLine("<!DOCTYPE html>");
                writer.WriteLine("<html lang=\"en\">");
                writer.WriteLine("<head>");
                writer.WriteLine("    <meta charset=\"UTF-8\">");
                writer.WriteLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
                writer.WriteLine("    <title>Image Gallery</title>");
                writer.WriteLine($"    <link rel=\"stylesheet\" href=\"{lightboxCss}\">");
                writer.WriteLine("    <style>");
                writer.WriteLine("        .gallery img {");
                writer.WriteLine("            width: 150px;");
                writer.WriteLine("            height: auto;");
                writer.WriteLine("            margin: 5px;");
                writer.WriteLine("            border: 2px solid #ccc;");
                writer.WriteLine("        }");
                writer.WriteLine("    </style>");
                writer.WriteLine("</head>");
                writer.WriteLine("<body>");
                writer.WriteLine("    <h1>Extracted Images Gallery</h1>");
                writer.WriteLine("    <div class=\"gallery\">");

                for (int i = 0; i < imagePaths.Count; i++)
                {
                    string imgPath = imagePaths[i];
                    writer.WriteLine($"        <a href=\"{imgPath}\" data-lightbox=\"gallery\" data-title=\"Image {i + 1}\">");
                    writer.WriteLine($"            <img src=\"{imgPath}\" alt=\"Image {i + 1}\">");
                    writer.WriteLine("        </a>");
                }

                writer.WriteLine("    </div>");
                writer.WriteLine($"    <script src=\"{lightboxJs}\"></script>");
                writer.WriteLine("</body>");
                writer.WriteLine("</html>");
            }

            // Validate HTML file creation
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to write HTML gallery to '{htmlPath}'.");
        }
    }
}
