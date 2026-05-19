using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtraction");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string htmlIndexPath = Path.Combine(baseDir, "index.html");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (100x100).
        // -------------------------------------------------
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Draw a simple rectangle.
        using (Pen pen = new Pen(Color.DarkBlue, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, 80, 80);
        }
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the image.
        // -------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i} with an embedded image.");
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputDocsDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process each DOCX: extract images.
        // -------------------------------------------------
        var docToImagesMap = new Dictionary<string, List<string>>();

        foreach (string docFile in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageCounter = 0;
            var extractedForDoc = new List<string>();

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageCounter}{extension}";
                string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);

                shape.ImageData.Save(imageFullPath);

                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save extracted image: {imageFullPath}");

                extractedForDoc.Add(imageFileName);
                imageCounter++;
            }

            if (imageCounter == 0)
                throw new InvalidOperationException($"No images were extracted from document: {docFile}");

            docToImagesMap[Path.GetFileName(docFile)] = extractedForDoc;
        }

        // -------------------------------------------------
        // 4. Generate a simple HTML index page.
        // -------------------------------------------------
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<title>Extracted Images Index</title>");
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: Arial, sans-serif; }");
        sb.AppendLine(".doc-section { margin-bottom: 30px; }");
        sb.AppendLine(".doc-section img { max-width: 200px; margin: 5px; border: 1px solid #ccc; }");
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<h1>Extracted Images Index</h1>");

        foreach (var kvp in docToImagesMap)
        {
            sb.AppendLine("<div class=\"doc-section\">");
            sb.AppendLine($"<h2>{kvp.Key}</h2>");
            foreach (string imgFile in kvp.Value)
            {
                string imgRelativePath = $"ExtractedImages/{imgFile}";
                sb.AppendLine($"<img src=\"{imgRelativePath}\" alt=\"{imgFile}\" />");
            }
            sb.AppendLine("</div>");
        }

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        File.WriteAllText(htmlIndexPath, sb.ToString());

        if (!File.Exists(htmlIndexPath))
            throw new InvalidOperationException("HTML index file was not created.");

        // Optional: write a brief console summary.
        Console.WriteLine($"Processed {docToImagesMap.Count} documents.");
        Console.WriteLine($"Extracted images are stored in: {extractedImagesDir}");
        Console.WriteLine($"HTML index page created at: {htmlIndexPath}");
    }
}
