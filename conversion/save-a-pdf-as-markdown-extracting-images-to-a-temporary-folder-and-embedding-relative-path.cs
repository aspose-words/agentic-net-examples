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
        // Prepare output directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string imagesDir = Path.Combine(outputDir, "images");
        Directory.CreateDirectory(imagesDir);

        // Create a simple PNG image using Aspose.Drawing.
        string tempImagePath = Path.Combine(outputDir, "temp.png");
        using (Bitmap bmp = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Blue);
            }
            bmp.Save(tempImagePath, ImageFormat.Png);
        }

        // Build a sample document that contains the image and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample document with an image.");
        builder.InsertImage(tempImagePath);
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF and convert it to Markdown, extracting images to a folder.
        Document pdfDoc = new Document(pdfPath);
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesDir,          // Physical folder where images will be written.
            ImagesFolderAlias = "images"       // Relative path used in the Markdown file.
        };
        string markdownPath = Path.Combine(outputDir, "sample.md");
        pdfDoc.Save(markdownPath, mdOptions);

        // Validate that the Markdown file and extracted images exist.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        string[] extractedImages = Directory.GetFiles(imagesDir);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // Output the locations of the generated files.
        Console.WriteLine($"Markdown file saved to: {markdownPath}");
        Console.WriteLine("Extracted image files:");
        foreach (string img in extractedImages)
            Console.WriteLine($"  {img}");
    }
}
