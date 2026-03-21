using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToHtmlWithEmbeddedFonts
{
    static void Main()
    {
        // Determine output path in the current directory.
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "output.html");

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Configure HTML fixed save options to embed fonts, CSS, and images as Base64.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true,
            ExportEmbeddedCss = true,
            ExportEmbeddedImages = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, htmlOptions);

        Console.WriteLine($"HTML file saved to: {htmlPath}");
    }
}
