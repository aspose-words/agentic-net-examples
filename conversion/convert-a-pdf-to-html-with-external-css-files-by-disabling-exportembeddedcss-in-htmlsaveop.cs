using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string pdfPath = "sample.pdf";
        string htmlPath = "output.html";

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it as PDF (input for conversion).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document created for conversion.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Convert PDF to HTML with external CSS files.
        //    ExportEmbeddedCss = false creates a separate CSS file.
        // -----------------------------------------------------------------
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false // Disable embedding, create external CSS.
        };

        pdfDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Validate that HTML and CSS files were generated.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        // CSS is saved in a subfolder named after the HTML file (without extension).
        string cssFolder = Path.Combine(
            Path.GetDirectoryName(htmlPath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(htmlPath));

        string cssPath = Path.Combine(cssFolder, "styles.css");

        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        // Output the locations of the generated files.
        Console.WriteLine($"PDF created at: {Path.GetFullPath(pdfPath)}");
        Console.WriteLine($"HTML created at: {Path.GetFullPath(htmlPath)}");
        Console.WriteLine($"CSS created at: {Path.GetFullPath(cssPath)}");
    }
}
