using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with embedded fonts.");

        // Configure HtmlSaveOptions to embed fonts as Base64.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,          // Embed fonts in Base64.
            ExportFontResources = true,          // Ensure font resources are exported.
            CssStyleSheetType = CssStyleSheetType.Embedded,
            PrettyFormat = true
        };

        // Save the document as HTML.
        const string outputFile = "output.html";
        doc.Save(outputFile, options);

        // Verify that the HTML file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Verify that the HTML contains Base64‑encoded font data.
        string htmlContent = File.ReadAllText(outputFile);
        if (!htmlContent.Contains("base64"))
            throw new InvalidOperationException("Base64 font data was not found in the HTML output.");

        Console.WriteLine("HTML file saved with embedded Base64 fonts.");
    }
}
