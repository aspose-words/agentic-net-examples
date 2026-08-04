using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string imagesDir = Path.Combine(workDir, "Images");
        string markdownImagesDir = Path.Combine(workDir, "MarkdownImages");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(markdownImagesDir);

        // Create a sample image using Aspose.Drawing (no System.Drawing)
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        using (Bitmap bmp = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.CornflowerBlue);
            }
            bmp.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document that contains text and the image
        string sourceDocxPath = Path.Combine(workDir, "Sample.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document with an image:");
        builder.InsertImage(sampleImagePath);
        sourceDoc.Save(sourceDocxPath, SaveFormat.Docx);

        // Load the DOCX document
        Document doc = new Document(sourceDocxPath);

        // Configure Markdown save options to extract images to a custom folder
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = markdownImagesDir,
            SaveFormat = SaveFormat.Markdown
        };

        // Save as Markdown
        string markdownPath = Path.Combine(workDir, "Sample.md");
        doc.Save(markdownPath, mdOptions);

        // Validation
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        if (!Directory.Exists(markdownImagesDir))
            throw new InvalidOperationException("Images folder was not created.");

        string[] extractedImages = Directory.GetFiles(markdownImagesDir);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted during Markdown conversion.");

        // Output paths (optional, not interactive)
        Console.WriteLine($"Markdown file: {markdownPath}");
        Console.WriteLine($"Extracted images folder: {markdownImagesDir}");
        foreach (string img in extractedImages)
        {
            Console.WriteLine($"Image: {img}");
        }
    }
}
