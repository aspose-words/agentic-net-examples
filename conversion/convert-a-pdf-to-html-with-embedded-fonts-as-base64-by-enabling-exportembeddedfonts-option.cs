using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the intermediate PDF and final HTML.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        string htmlPath = Path.Combine(outputDir, "sample.html");

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it as PDF (the source PDF).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello PDF to HTML conversion with embedded fonts.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // --------------------------------------------------------------
        // 2. Load the PDF we just created.
        // --------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // --------------------------------------------------------------
        // 3. Configure HTML Fixed save options to embed fonts as Base64.
        // --------------------------------------------------------------
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true,   // Embed fonts in Base64.
            ExportEmbeddedCss = true,    // Keep CSS in the same file (optional).
            PrettyFormat = true
        };

        // --------------------------------------------------------------
        // 4. Save the PDF as HTML using the configured options.
        // --------------------------------------------------------------
        pdfDoc.Save(htmlPath, htmlOptions);

        // --------------------------------------------------------------
        // 5. Validate that the HTML file exists and contains Base64 data.
        // --------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new FileNotFoundException("HTML output file was not created.", htmlPath);

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("base64"))
            throw new InvalidOperationException("The HTML does not contain Base64‑encoded fonts.");

        // Indicate successful completion.
        Console.WriteLine("Conversion completed successfully.");
    }
}
