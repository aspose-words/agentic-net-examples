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
        // Define folders for input documents, extracted images and the HTML index.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string htmlPath = Path.Combine(baseDir, "ImageIndex.html");

        // Ensure clean environment.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(imagesDir)) Directory.Delete(imagesDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);

        // ------------------------------------------------------------
        // 1. Create sample images that will be inserted into the documents.
        // ------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 150, Color.LightBlue);

        // ------------------------------------------------------------
        // 2. Create a few sample DOCX files, each containing the sample image.
        // ------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {docIndex}");
            // Insert the same image twice to have multiple images per document.
            builder.InsertImage(sampleImagePath);
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath);

            string docPath = Path.Combine(inputDir, $"SampleDoc{docIndex}.docx");
            doc.Save(docPath);
        }

        // ------------------------------------------------------------
        // 3. Batch process all DOCX files: extract images and build HTML.
        // ------------------------------------------------------------
        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html><head><meta charset=\"UTF-8\"><title>Extracted Images Index</title></head><body>");
        htmlBuilder.AppendLine("<h1>Extracted Images</h1>");

        int totalExtracted = 0;

        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image to the output folder.
                shape.ImageData.Save(imageFullPath);
                imageIndex++;
                totalExtracted++;

                // Add entry to HTML.
                htmlBuilder.AppendLine("<div style=\"margin-bottom:20px;\">");
                htmlBuilder.AppendLine($"<p>Document: {Path.GetFileName(docFile)}</p>");
                htmlBuilder.AppendLine($"<img src=\"{Path.GetFileName(imageFullPath)}\" alt=\"{imageFileName}\" style=\"max-width:600px;\"/>");
                htmlBuilder.AppendLine("</div>");
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        htmlBuilder.AppendLine("</body></html>");

        // Save the HTML index file (images are referenced relative to the HTML file location).
        File.WriteAllText(htmlPath, htmlBuilder.ToString());

        // ------------------------------------------------------------
        // 4. Clean up temporary sample image.
        // ------------------------------------------------------------
        if (File.Exists(sampleImagePath))
            File.Delete(sampleImagePath);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height, Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
                // Simple drawing: a diagonal line.
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawLine(pen, 0, 0, width, height);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
