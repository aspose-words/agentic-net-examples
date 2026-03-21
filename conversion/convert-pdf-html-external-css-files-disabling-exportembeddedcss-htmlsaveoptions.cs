using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToHtmlWithExternalCss
{
    static void Main()
    {
        // Input PDF path (optional). If the file does not exist, create a simple document.
        string pdfPath = "source.pdf";
        Document doc;

        if (System.IO.File.Exists(pdfPath))
        {
            doc = new Document(pdfPath);
        }
        else
        {
            // Create a simple document as a placeholder.
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document generated because the source PDF was not found.");
        }

        // Output HTML path.
        string htmlPath = "output.html";

        // Configure HTML fixed save options to export external CSS.
        var htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false // false creates an external CSS file
        };

        // Save the document as HTML.
        doc.Save(htmlPath, htmlOptions);
        Console.WriteLine($"HTML saved to {htmlPath}");
    }
}
