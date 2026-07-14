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
        // Base output directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(baseDir);

        // -----------------------------------------------------------------
        // 1. Create a simple image using Aspose.Drawing (no System.Drawing).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a PDF document that contains some text and the image.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content with an embedded image:");
        builder.InsertImage(sampleImagePath);

        string pdfPath = Path.Combine(baseDir, "sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF and convert it to Markdown, extracting images.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Folder where extracted images will be saved.
        string imagesFolder = Path.Combine(baseDir, "markdown_images");
        Directory.CreateDirectory(imagesFolder);

        // Markdown file path.
        string markdownPath = Path.Combine(baseDir, "sample.md");

        // Configure Markdown save options.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,          // Physical folder for image files.
            ImagesFolderAlias = "images",         // Relative path used inside the .md file.
            SaveFormat = SaveFormat.Markdown     // Explicitly set the format.
        };

        // Perform the conversion.
        pdfDoc.Save(markdownPath, mdOptions);

        // -----------------------------------------------------------------
        // 4. Validation – ensure the Markdown file and extracted images exist.
        // -----------------------------------------------------------------
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted during conversion.");

        // (Optional) Output the paths for verification – no interactive prompts.
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine("Markdown file: " + markdownPath);
        Console.WriteLine("Extracted images folder: " + imagesFolder);
        Console.WriteLine("Number of images extracted: " + extractedImages.Length);
    }
}
