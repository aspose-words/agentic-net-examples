using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a simple PDF document that will be used as input.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);
        builder.Writeln("This is a sample PDF document created for conversion.");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // ---------------------------------------------------------------
        // Step 2: Load the PDF and convert it to HTML with external CSS.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // HtmlFixedSaveOptions is used because it provides the ExportEmbeddedCss property.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            // When false, CSS is saved to an external file instead of being embedded.
            ExportEmbeddedCss = false
        };

        string htmlPath = Path.Combine(outputDir, "sample.html");
        pdfDoc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML conversion failed: HTML file not found.");

        // The external CSS file is placed in a subfolder named after the HTML file (without extension).
        string cssFolder = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(htmlPath));
        string cssPath = Path.Combine(cssFolder, "styles.css");

        // Verify that the external CSS file exists.
        if (!File.Exists(cssPath))
            throw new InvalidOperationException("HTML conversion failed: external CSS file not found.");

        // Output the locations of the generated files (optional, not required for the task).
        Console.WriteLine($"PDF created at: {pdfPath}");
        Console.WriteLine($"HTML created at: {htmlPath}");
        Console.WriteLine($"External CSS created at: {cssPath}");
    }
}
