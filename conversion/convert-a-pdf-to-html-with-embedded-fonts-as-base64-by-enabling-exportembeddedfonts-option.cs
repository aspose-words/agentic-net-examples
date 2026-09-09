using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string pdfPath = "sample.pdf";
        string htmlPath = "sample.html";

        // 1. Create a simple document and save it as PDF
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample PDF document with embedded fonts.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF was created
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF file was not created.");

        // 2. Load the PDF document
        Document pdfDoc = new Document(pdfPath);

        // 3. Configure HtmlFixedSaveOptions to embed fonts as Base64
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true
        };

        // 4. Save the PDF as HTML with embedded fonts
        pdfDoc.Save(htmlPath, htmlOptions);

        // 5. Validate that the HTML output exists and contains data
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            throw new InvalidOperationException("HTML file was not created or is empty.");

        // Optional: Output a simple confirmation (no interactive input required)
        Console.WriteLine("Conversion completed successfully.");
    }
}
