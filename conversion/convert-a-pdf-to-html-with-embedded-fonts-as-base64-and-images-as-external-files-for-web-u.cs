using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, etc.

public class PdfToHtmlConverter
{
    public static void Main()
    {
        // Base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ConversionDemo");
        Directory.CreateDirectory(baseDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(baseDir, "sample.pdf");
        CreateSamplePdf(pdfPath, baseDir);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to HTML.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);

        // Configure HTML save options:
        // - Embed fonts as Base64.
        // - Keep images as external files.
        // - Specify a folder for the external images.
        string imagesFolder = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imagesFolder);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,
            ExportImagesAsBase64 = false,
            ImagesFolder = imagesFolder,
            PrettyFormat = true
        };

        string htmlPath = Path.Combine(baseDir, "sample.html");
        pdfDocument.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // Step 3: Validation.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new FileNotFoundException("HTML output file was not created.", htmlPath);

        // Verify that at least one image file was written to the images folder.
        string[] imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No image files were generated in the Images folder.");

        // Output the locations of the generated files.
        Console.WriteLine($"PDF created at: {pdfPath}");
        Console.WriteLine($"HTML created at: {htmlPath}");
        Console.WriteLine($"Images folder contains {imageFiles.Length} file(s).");
    }

    private static void CreateSamplePdf(string pdfPath, string baseDir)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("This is a sample PDF document generated for conversion.");

        // Create a tiny PNG image (1x1 pixel) using Aspose.Drawing.
        string pngPath = Path.Combine(baseDir, "sample.png");
        CreateTinyPng(pngPath);

        // Insert the image into the document.
        builder.InsertImage(pngPath);

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);
    }

    private static void CreateTinyPng(string filePath)
    {
        // 1x1 pixel transparent PNG (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }
}
