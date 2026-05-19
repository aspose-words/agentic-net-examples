using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document and save it as PDF – this will be our input file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion to HTML with embedded fonts.");
        const string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Configure HTML fixed save options to embed fonts as Base64.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true
        };

        // Save the PDF as HTML using the configured options.
        const string htmlPath = "output.html";
        pdfDocument.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("PDF successfully converted to HTML with embedded fonts.");
    }
}
