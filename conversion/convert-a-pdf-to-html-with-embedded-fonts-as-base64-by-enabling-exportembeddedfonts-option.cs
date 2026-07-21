using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial"; // Use a common font to demonstrate embedding.
        builder.Writeln("Sample PDF content with embedded fonts.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Step 2: Load the generated PDF.
        Document pdfDoc = new Document(pdfPath);

        // Step 3: Configure HtmlFixedSaveOptions to embed fonts as Base64.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedFonts = true,   // Embed fonts in Base64.
            ExportEmbeddedCss = true,    // Optional: embed CSS to keep a single file.
            PrettyFormat = true
        };

        // Step 4: Save the PDF as HTML using the configured options.
        string htmlPath = "output.html";
        pdfDoc.Save(htmlPath, htmlOptions);

        // Step 5: Validate that the HTML file was created and is not empty.
        if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
        {
            throw new InvalidOperationException("HTML conversion failed: output file is missing or empty.");
        }
    }
}
