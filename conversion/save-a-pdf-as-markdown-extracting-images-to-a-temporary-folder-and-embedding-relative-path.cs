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
        // Define base output directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(baseDir);

        // Paths for temporary files.
        string imagePath = Path.Combine(baseDir, "sample.png");
        string pdfPath = Path.Combine(baseDir, "sample.pdf");
        string markdownPath = Path.Combine(baseDir, "sample.md");
        string imagesFolder = Path.Combine(baseDir, "md_images");

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image using Aspose.Drawing (no System.Drawing).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document, insert the image, and save it as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image that will be extracted when converting to Markdown.");
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF back into a Document object.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 4. Configure MarkdownSaveOptions to extract images to a folder.
        // -----------------------------------------------------------------
        Directory.CreateDirectory(imagesFolder);
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,          // Physical folder where images will be saved.
            ImagesFolderAlias = "images",         // Relative path used inside the Markdown file.
            SaveFormat = SaveFormat.Markdown    // Explicitly set the format.
        };

        // -----------------------------------------------------------------
        // 5. Save the document as Markdown.
        // -----------------------------------------------------------------
        pdfDoc.Save(markdownPath, mdOptions);

        // -----------------------------------------------------------------
        // 6. Validation: ensure Markdown file exists and images were extracted.
        // -----------------------------------------------------------------
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        string[] extractedImages = Directory.GetFiles(imagesFolder);
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // -----------------------------------------------------------------
        // 7. Output simple confirmation (no user interaction required).
        // -----------------------------------------------------------------
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine($"Markdown file: {markdownPath}");
        Console.WriteLine($"Extracted images folder: {imagesFolder}");
        Console.WriteLine($"Number of extracted images: {extractedImages.Length}");
    }
}
