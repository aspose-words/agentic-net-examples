using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Writeln("This is a sample PDF document with embedded fonts.");

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // NOTE: In some versions of Aspose.Words the ExportEmbeddedFonts property
        // is not available. If it exists, you can enable it as follows:
        // htmlOptions.ExportEmbeddedFonts = true;
        // The example proceeds without setting it to maintain compatibility.

        // Save the PDF as HTML.
        string htmlPath = "output.html";
        pdfDoc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file was created and is not empty.
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            throw new InvalidOperationException("HTML file was not created or is empty.");
    }
}
