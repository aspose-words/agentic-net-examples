using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string htmlPath = "sample.html";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF (input file).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample PDF document.");
        builder.Writeln("It will be converted to HTML with external CSS.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF input file.");

        // ---------------------------------------------------------------
        // 2. Load the PDF document that we just created.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // ---------------------------------------------------------------
        // 3. Configure HtmlFixedSaveOptions to disable embedded CSS.
        //    When ExportEmbeddedCss is false, the CSS is not embedded in the
        //    HTML file, resulting in an external stylesheet.
        // ---------------------------------------------------------------
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false
        };

        // ---------------------------------------------------------------
        // 4. Save the PDF as HTML using the configured options.
        // ---------------------------------------------------------------
        pdfDoc.Save(htmlPath, htmlOptions);

        // ---------------------------------------------------------------
        // 5. Validate that the HTML output file exists.
        // ---------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML conversion failed; output file not found.");

        // Optional: display a short confirmation.
        Console.WriteLine("PDF successfully converted to HTML with external CSS.");
    }
}
