using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(baseDir);

        // Paths for the sample PDF, resulting HTML, CSS file and images folder.
        string pdfPath = Path.Combine(baseDir, "sample.pdf");
        string htmlPath = Path.Combine(baseDir, "sample.html");
        string cssPath = Path.Combine(baseDir, "styles.css");
        string imagesFolder = Path.Combine(baseDir, "Images");

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it as PDF (input for conversion).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words! This PDF will be converted to HTML.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and configure HTML save options.
        //    - External CSS file.
        //    - Separate folder for images.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Ensure the images folder exists before saving.
        Directory.CreateDirectory(imagesFolder);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External, // External CSS.
            CssStyleSheetFileName = cssPath,                // Path for the CSS file.
            ImagesFolder = imagesFolder,                    // Folder where images will be saved.
            ExportFontResources = false,                    // No font export needed for this example.
            ExportOriginalUrlForLinkedImages = false        // Use local image files.
        };

        // -----------------------------------------------------------------
        // 3. Perform the conversion.
        // -----------------------------------------------------------------
        pdfDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Validation – ensure all expected outputs exist.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        // Optional: verify that the images folder contains at least one file when the source PDF has images.
        // For this simple text‑only PDF the folder may be empty, which is acceptable.

        Console.WriteLine("PDF successfully converted to HTML.");
        Console.WriteLine($"HTML file: {htmlPath}");
        Console.WriteLine($"CSS file: {cssPath}");
        Console.WriteLine($"Images folder: {imagesFolder}");
    }
}
