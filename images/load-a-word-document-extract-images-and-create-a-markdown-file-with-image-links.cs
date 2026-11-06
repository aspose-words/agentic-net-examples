using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, sampleImagePath);

        // Step 3: Load the document.
        Document doc = new Document(docPath);

        // Step 4: Extract images from the document.
        List<string> extractedImagePaths = ExtractImages(doc);

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Step 5: Generate a Markdown file with image links.
        const string markdownPath = "images.md";
        GenerateMarkdown(markdownPath, extractedImagePaths);
    }

    private static void CreateSampleImage(string path)
    {
        // Create a 200x200 white bitmap.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle for visual distinction.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
                {
                    graphics.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(path, ImageFormat.Png);
        }

        // Ensure the file exists.
        if (!File.Exists(path))
            throw new FileNotFoundException($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image into the document.
        builder.InsertImage(imagePath);
        // Save the document.
        doc.Save(docPath);
        // Validate the document was saved.
        if (!File.Exists(docPath))
            throw new FileNotFoundException($"Failed to save document at '{docPath}'.");
    }

    private static List<string> ExtractImages(Document doc)
    {
        List<string> imagePaths = new List<string>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 1;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"extracted-{index}.png";
                shape.ImageData.Save(imageFileName);
                if (!File.Exists(imageFileName))
                    throw new FileNotFoundException($"Failed to save extracted image '{imageFileName}'.");
                imagePaths.Add(imageFileName);
                index++;
            }
        }
        return imagePaths;
    }

    private static void GenerateMarkdown(string markdownPath, List<string> imagePaths)
    {
        using (StreamWriter writer = new StreamWriter(markdownPath, false))
        {
            foreach (string imgPath in imagePaths)
            {
                writer.WriteLine($"![]({imgPath})");
            }
        }

        // Validate markdown file creation.
        if (!File.Exists(markdownPath))
            throw new FileNotFoundException($"Failed to create markdown file at '{markdownPath}'.");
    }
}
