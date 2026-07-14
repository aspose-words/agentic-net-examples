using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        const string pdfPath = "sample.pdf";
        const string htmlPath = "output.html";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF (input file).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF document.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // ---------------------------------------------------------------
        // 2. Load the PDF and convert it to HTML with embedded fonts (Base64).
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true   // Embed fonts as Base64 in the CSS.
        };

        pdfDoc.Save(htmlPath, htmlOptions);

        // Validate that the HTML output was created and contains data.
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            throw new InvalidOperationException("HTML conversion failed; output file is missing or empty.");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine("PDF successfully converted to HTML with embedded fonts.");
    }
}
