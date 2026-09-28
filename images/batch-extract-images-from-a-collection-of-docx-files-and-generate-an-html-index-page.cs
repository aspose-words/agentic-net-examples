using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Define folders
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string outputFolder = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files containing the sample image
        CreateSampleDocument(Path.Combine(inputFolder, "doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "doc2.docx"), sampleImagePath);

        // List to hold extracted image relative paths for HTML generation
        List<string> extractedImageRelativePaths = new List<string>();

        // Process each DOCX file in the input folder
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.png";
                    string imageFullPath = Path.Combine(imageFolder, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    // Store relative path for HTML (relative to output folder)
                    string relativePath = Path.Combine("..", "ExtractedImages", imageFileName).Replace('\\', '/');
                    extractedImageRelativePaths.Add(relativePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImageRelativePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Generate HTML index page
        string htmlPath = Path.Combine(outputFolder, "index.html");
        GenerateHtmlIndex(htmlPath, extractedImageRelativePaths);

        // Validate HTML file creation
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Failed to create the HTML index page.");
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Optional: draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }

        // Validate that the image file exists
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        // Validate that the document file exists
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create sample document at '{docPath}'.");
    }

    private static void GenerateHtmlIndex(string htmlPath, List<string> imagePaths)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head><meta charset=\"UTF-8\"><title>Extracted Images Index</title></head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<h1>Extracted Images</h1>");
        foreach (string relPath in imagePaths)
        {
            sb.AppendLine($"<div><img src=\"{relPath}\" alt=\"Extracted Image\" style=\"max-width:300px;\"/></div>");
        }
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        File.WriteAllText(htmlPath, sb.ToString());
    }
}
