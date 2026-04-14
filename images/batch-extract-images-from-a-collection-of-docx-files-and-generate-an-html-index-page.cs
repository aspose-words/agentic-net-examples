using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing namespace for image creation

public class Program
{
    public static void Main()
    {
        // Base working directory for the demo.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "DemoData");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure all required folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images (no external assets).
        // -----------------------------------------------------------------
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(sampleImage1, 200, 150, Color.LightBlue);
        CreateSampleImage(sampleImage2, 150, 200, Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -----------------------------------------------------------------
        string doc1Path = Path.Combine(inputDir, "Document1.docx");
        string doc2Path = Path.Combine(inputDir, "Document2.docx");

        CreateSampleDocument(doc1Path, new[] { sampleImage1, sampleImage2 });
        CreateSampleDocument(doc2Path, new[] { sampleImage2 });

        // -----------------------------------------------------------------
        // 3. Batch extract images from all DOCX files in the input folder.
        // -----------------------------------------------------------------
        var extractedInfo = new List<DocumentImagesInfo>();

        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            var docInfo = ExtractImagesFromDocument(docPath, imagesDir);
            extractedInfo.Add(docInfo);
        }

        // Validate that at least one image was extracted overall.
        if (!extractedInfo.Any(info => info.ImageFiles.Any()))
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -----------------------------------------------------------------
        // 4. Generate an HTML index page that lists each document and its images.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(outputDir, "index.html");
        GenerateHtmlIndex(htmlPath, extractedInfo, imagesDir, outputDir);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Failed to create the HTML index page.");
    }

    // -----------------------------------------------------------------
    // Helper: creates a deterministic PNG image using Aspose.Drawing.
    // -----------------------------------------------------------------
    private static void CreateSampleImage(string filePath, int width, int height, Color background)
    {
        // Ensure any existing file is overwritten.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and draw background.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(background);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Verify that the image file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    // -----------------------------------------------------------------
    // Helper: creates a DOCX file and inserts the provided images.
    // -----------------------------------------------------------------
    private static void CreateSampleDocument(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln($"Document generated on {DateTime.Now}");

        foreach (string imgPath in imagePaths)
        {
            // Insert the image using the builder (inline shape).
            Shape shape = builder.InsertImage(imgPath);
            // Ensure the shape actually contains an image.
            if (!shape.HasImage)
                throw new InvalidOperationException($"Failed to insert image '{imgPath}' into document.");
            builder.Writeln(); // Add a line break after each image.
        }

        doc.Save(docPath);

        // Verify that the document file exists.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document at '{docPath}'.");
    }

    // -----------------------------------------------------------------
    // Helper: extracts all images from a single document.
    // Returns information needed for HTML generation.
    // -----------------------------------------------------------------
    private static DocumentImagesInfo ExtractImagesFromDocument(string docPath, string outputImagesDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        var imageFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on the image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputImagesDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);
                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save extracted image '{imageFullPath}'.");

                imageFiles.Add(imageFullPath);
                imageIndex++;
            }
        }

        return new DocumentImagesInfo
        {
            DocumentPath = docPath,
            ImageFiles = imageFiles
        };
    }

    // -----------------------------------------------------------------
    // Helper: generates a simple HTML index page linking to extracted images.
    // -----------------------------------------------------------------
    private static void GenerateHtmlIndex(string htmlFilePath, List<DocumentImagesInfo> docsInfo, string imagesFolder, string outputFolder)
    {
        // Compute relative path from HTML file to images folder.
        string relativeImagesPath = GetRelativePath(outputFolder, imagesFolder).Replace('\\', '/');

        using (StreamWriter writer = new StreamWriter(htmlFilePath, false))
        {
            writer.WriteLine("<!DOCTYPE html>");
            writer.WriteLine("<html lang=\"en\">");
            writer.WriteLine("<head><meta charset=\"UTF-8\"><title>Extracted Images Index</title></head>");
            writer.WriteLine("<body>");
            writer.WriteLine("<h1>Extracted Images Index</h1>");

            foreach (var docInfo in docsInfo)
            {
                writer.WriteLine($"<h2>{Path.GetFileName(docInfo.DocumentPath)}</h2>");

                if (docInfo.ImageFiles.Any())
                {
                    foreach (string imgPath in docInfo.ImageFiles)
                    {
                        string relativeImgPath = Path.Combine(relativeImagesPath, Path.GetFileName(imgPath)).Replace('\\', '/');
                        writer.WriteLine($"<div style=\"margin:10px 0;\"><img src=\"{relativeImgPath}\" alt=\"{Path.GetFileName(imgPath)}\" style=\"max-width:400px; height:auto;\"/></div>");
                    }
                }
                else
                {
                    writer.WriteLine("<p>No images found in this document.</p>");
                }
            }

            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
        }
    }

    // -----------------------------------------------------------------
    // Utility: computes a relative path from one folder to another.
    // -----------------------------------------------------------------
    private static string GetRelativePath(string fromPath, string toPath)
    {
        Uri fromUri = new Uri(AppendDirectorySeparatorChar(fromPath));
        Uri toUri = new Uri(AppendDirectorySeparatorChar(toPath));

        Uri relativeUri = fromUri.MakeRelativeUri(toUri);
        string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

        return relativePath.Replace('/', Path.DirectorySeparatorChar);
    }

    private static string AppendDirectorySeparatorChar(string path)
    {
        if (!path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            return path + Path.DirectorySeparatorChar;
        return path;
    }

    // -----------------------------------------------------------------
    // Simple DTO to hold extraction results per document.
    // -----------------------------------------------------------------
    private class DocumentImagesInfo
    {
        public string DocumentPath { get; set; }
        public List<string> ImageFiles { get; set; } = new List<string>();
    }
}
