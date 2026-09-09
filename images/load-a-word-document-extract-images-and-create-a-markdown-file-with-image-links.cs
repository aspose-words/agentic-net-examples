using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define folders for output.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string imagesDir = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesDir);

        // 1. Create a deterministic sample image using Aspose.Drawing.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 150);

        // 2. Create a Word document and insert the sample image.
        string docPath = Path.Combine(baseDir, "sample.docx");
        CreateWordDocumentWithImage(docPath, sampleImagePath);

        // 3. Load the document.
        Document doc = new Document(docPath);

        // 4. Extract images from the document.
        List<string> extractedImageFiles = ExtractImages(doc, imagesDir);

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 5. Generate a Markdown file with image links.
        string markdownPath = Path.Combine(baseDir, "document.md");
        GenerateMarkdownFile(markdownPath, extractedImageFiles, "Images");

        // Indicate successful completion.
        Console.WriteLine("Markdown file created at: " + markdownPath);
    }

    // Creates a simple PNG image with a solid background.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                // Additional deterministic drawing can be added here if needed.
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a Word document and inserts the specified image.
    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts all images from the document and saves them to the target folder.
    private static List<string> ExtractImages(Document doc, string targetFolder)
    {
        List<string> savedFiles = new List<string>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"img{imageIndex}{extension}";
                string fullPath = Path.Combine(targetFolder, fileName);
                shape.ImageData.Save(fullPath);
                savedFiles.Add(fileName);
                imageIndex++;
            }
        }

        return savedFiles;
    }

    // Generates a Markdown file that references the extracted images.
    private static void GenerateMarkdownFile(string markdownPath, List<string> imageFiles, string imagesFolderAlias)
    {
        using (StreamWriter writer = new StreamWriter(markdownPath, false))
        {
            writer.WriteLine("# Extracted Images");
            writer.WriteLine();

            for (int i = 0; i < imageFiles.Count; i++)
            {
                string relativePath = Path.Combine(imagesFolderAlias, imageFiles[i]).Replace('\\', '/');
                writer.WriteLine($"![Image {i}]({relativePath})");
                writer.WriteLine();
            }
        }
    }
}
