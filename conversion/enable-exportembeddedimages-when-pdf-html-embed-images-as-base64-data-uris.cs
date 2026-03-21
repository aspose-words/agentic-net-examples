using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a temporary PDF file.
        string tempPdf = Path.Combine(Path.GetTempPath(), "sample.pdf");
        if (File.Exists(tempPdf))
            File.Delete(tempPdf);

        // Create a simple Word document and save it as PDF.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Writeln("This is a sample PDF generated for conversion.");
        tempDoc.Save(tempPdf, SaveFormat.Pdf);

        // Path where the resulting HTML file will be saved.
        string htmlPath = Path.Combine(Path.GetTempPath(), "sample.html");

        // Load the PDF document.
        Document doc = new Document(tempPdf);

        // Configure HTML Fixed save options to embed images as Base64 data URIs.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedImages = true,
            PrettyFormat = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, htmlOptions);

        Console.WriteLine($"PDF has been converted to HTML with embedded images.\nHTML saved to: {htmlPath}");
    }
}
