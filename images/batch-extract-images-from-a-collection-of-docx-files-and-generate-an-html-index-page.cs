using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files that contain the image
        CreateSampleDocuments(inputFolder, sampleImagePath, 3);

        // Batch extract images from all DOCX files
        var extractedFiles = ExtractImagesFromDocuments(inputFolder, imageFolder);

        // Validate that at least one image was extracted
        if (extractedFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Generate HTML index page
        string htmlPath = Path.Combine(baseDir, "index.html");
        GenerateHtmlIndex(htmlPath, extractedFiles, imageFolder);

        // Validation of HTML output
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML index page was not created.");

        // Optional: write a short console summary
        Console.WriteLine($"Processed {extractedFiles.Count} images. HTML index created at: {htmlPath}");
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        // Ensure any existing file is removed
        if (File.Exists(path))
            File.Delete(path);

        // Create deterministic bitmap using Aspose.Drawing
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        // Draw a simple rectangle for visual distinction
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
        {
            g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }
        bitmap.Save(path);
        g.Dispose();
        bitmap.Dispose();
    }

    private static void CreateSampleDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertImage(imagePath);
            string docPath = Path.Combine(folder, $"Doc{i}.docx");
            doc.Save(docPath);
        }
    }

    private static List<string> ExtractImagesFromDocuments(string inputFolder, string outputFolder)
    {
        var extracted = new List<string>();
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(Aspose.Words.NodeType.Shape, true);
            int imgIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imgFileName = $"{Path.GetFileNameWithoutExtension(docPath)}-img{imgIndex}.png";
                    string imgFullPath = Path.Combine(outputFolder, imgFileName);
                    shape.ImageData.Save(imgFullPath);
                    extracted.Add(imgFileName);
                    imgIndex++;
                }
            }
        }
        return extracted;
    }

    private static void GenerateHtmlIndex(string htmlPath, List<string> imageFiles, string imageFolder)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head><meta charset=\"UTF-8\"><title>Image Index</title></head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<h1>Extracted Images</h1>");
        foreach (string imgFile in imageFiles)
        {
            string relativePath = Path.Combine("ExtractedImages", imgFile).Replace("\\", "/");
            sb.AppendLine("<div style=\"margin-bottom:20px;\">");
            sb.AppendLine($"<p>{imgFile}</p>");
            sb.AppendLine($"<img src=\"{relativePath}\" alt=\"{imgFile}\" style=\"max-width:400px;\"/>");
            sb.AppendLine("</div>");
        }
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        File.WriteAllText(htmlPath, sb.ToString());
    }
}
