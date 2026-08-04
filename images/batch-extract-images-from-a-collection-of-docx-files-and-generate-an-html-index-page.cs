using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Define folders for input documents, extracted images and the HTML index.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image that will be inserted into the sample documents.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a few sample DOCX files containing the sample image.
        CreateSampleDocuments(inputDir, sampleImagePath, 3);

        // Dictionary to hold extracted image file names per document.
        var docImagesMap = new Dictionary<string, List<string>>();

        int totalExtracted = 0;

        // Process each DOCX file in the input directory.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            var extractedImages = ExtractImagesFromDocument(docPath, imagesDir);
            if (extractedImages.Count > 0)
            {
                docImagesMap[Path.GetFileName(docPath)] = extractedImages;
                totalExtracted += extractedImages.Count;
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Generate an HTML index page linking to the extracted images.
        string htmlPath = Path.Combine(outputDir, "index.html");
        GenerateHtmlIndex(htmlPath, docImagesMap, imagesDir);
    }

    // Creates a simple white PNG image with a black rectangle.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple black rectangle for visual distinction.
                graphics.DrawRectangle(new Pen(Color.Black, 5), 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath);
        }
    }

    // Generates a number of DOCX files each containing the sample image.
    private static void CreateSampleDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i}");
            // Insert the deterministic image.
            builder.InsertImage(imagePath);
            string docPath = Path.Combine(folder, $"Document{i}.docx");
            doc.Save(docPath);
        }
    }

    // Extracts all images from a single document and saves them to the target folder.
    private static List<string> ExtractImagesFromDocument(string docPath, string targetFolder)
    {
        var savedFiles = new List<string>();
        Document doc = new Document(docPath);

        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string fullPath = Path.Combine(targetFolder, imageFileName);
                shape.ImageData.Save(fullPath);
                savedFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        return savedFiles;
    }

    // Builds a simple HTML file that lists each document and its extracted images.
    private static void GenerateHtmlIndex(string htmlFilePath, Dictionary<string, List<string>> docImagesMap, string imagesFolder)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head><meta charset=\"UTF-8\"><title>Extracted Images Index</title></head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<h1>Extracted Images Index</h1>");

        foreach (var entry in docImagesMap)
        {
            sb.AppendLine($"<h2>{entry.Key}</h2>");
            foreach (string imageFile in entry.Value)
            {
                string relativePath = Path.Combine("..", "ExtractedImages", imageFile).Replace('\\', '/');
                sb.AppendLine($"<img src=\"{relativePath}\" alt=\"{imageFile}\" style=\"margin:5px;max-width:300px;\"/>");
            }
        }

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        File.WriteAllText(htmlFilePath, sb.ToString());
    }
}
