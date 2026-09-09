using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define directories for output artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "Images");

        // Ensure clean state.
        if (Directory.Exists(artifactsDir))
            Directory.Delete(artifactsDir, true);
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document and save it as PDF (input for conversion).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document generated for conversion.");
        // Save as PDF.
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure HTML save options:
        //    - External CSS file.
        //    - Separate folder for images.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(artifactsDir, "sample.html");
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External, // External CSS.
            ImagesFolder = imagesDir,                       // Folder for extracted images.
            // Optional: specify a custom CSS filename.
            // CssStyleSheetFileName = Path.Combine(artifactsDir, "sample.css")
        };

        // -----------------------------------------------------------------
        // 4. Convert PDF to HTML.
        // -----------------------------------------------------------------
        pdfDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validation of output files.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // CSS file should be created alongside the HTML file.
        string cssPath = Path.ChangeExtension(htmlPath, ".css");
        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        // Images folder should exist (may be empty if no images were present).
        if (!Directory.Exists(imagesDir))
            throw new InvalidOperationException("Images folder was not created.");

        // Example completed successfully.
        Console.WriteLine("PDF successfully converted to HTML with external CSS and image folder.");
        Console.WriteLine($"HTML file: {htmlPath}");
        Console.WriteLine($"CSS file: {cssPath}");
        Console.WriteLine($"Images folder: {imagesDir}");
    }
}
