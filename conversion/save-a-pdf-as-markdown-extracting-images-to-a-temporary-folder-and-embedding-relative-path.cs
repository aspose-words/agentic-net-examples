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
        // Prepare folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(baseDir);

        string imagesFolder = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesFolder);

        // Create a simple PNG image using Aspose.Drawing.
        string imagePath = Path.Combine(baseDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                graphics.DrawEllipse(new Pen(Color.DarkBlue, 3), 10, 10, 80, 80);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a source document, insert the image, and save it as PDF.
        string pdfPath = Path.Combine(baseDir, "input.pdf");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF containing an image.");
        builder.InsertImage(imagePath);
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure Markdown save options to extract images to a folder
        // and use a relative path (alias) in the generated Markdown.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = "images",
            SaveFormat = SaveFormat.Markdown
        };

        // Save the PDF as Markdown.
        string markdownPath = Path.Combine(baseDir, "output.md");
        pdfDoc.Save(markdownPath, mdOptions);

        // Validation.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted during conversion.");

        // Optional: display the result paths.
        Console.WriteLine("Markdown file created at: " + markdownPath);
        Console.WriteLine("Extracted images:");
        foreach (string img in extractedImages)
            Console.WriteLine(" - " + img);
    }
}
