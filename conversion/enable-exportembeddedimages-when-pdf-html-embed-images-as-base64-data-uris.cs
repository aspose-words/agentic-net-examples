using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path where the resulting HTML file will be saved.
        string htmlPath = @"C:\Output\sample.html";

        // Load the PDF document into an Aspose.Words Document object.
        Document doc = new Document(pdfPath);

        // Configure HTML Fixed save options to embed images as Base64 data URIs.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            // When true, images are stored directly in the HTML <img> tags as Base64 strings.
            ExportEmbeddedImages = true,

            // Optional: make the output more readable.
            PrettyFormat = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, htmlOptions);

        Console.WriteLine("PDF has been converted to HTML with embedded images.");
    }
}
