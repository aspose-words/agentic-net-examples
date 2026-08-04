using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample text for PDF to HTML conversion.");
        // Save the document as PDF – this will be the input file.
        const string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure HTML save options to embed fonts as Base64.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true,   // Embed fonts in Base64.
            ExportEmbeddedCss = true,    // Embed CSS to keep a single file.
            PrettyFormat = true
        };

        // Save the PDF as HTML with the specified options.
        const string htmlPath = "output.html";
        pdfDoc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created and contains Base64 data.
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            throw new InvalidOperationException("The HTML output file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("base64"))
            throw new InvalidOperationException("Fonts were not embedded as Base64 in the HTML output.");
    }
}
