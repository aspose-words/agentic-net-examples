using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a deterministic output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create a sample Word document in memory.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello, this is a sample document with embedded fonts.");

        // Save the DOCX just to demonstrate the input‑bootstrap rule (optional).
        string docPath = Path.Combine(outputFolder, "Sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Configure HTML save options to embed fonts as Base64.
        // ExportFontResources must be true for fonts to be exported,
        // and ExportFontsAsBase64 tells Aspose.Words to embed them in Base64.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportFontResources = true,          // Enable font exporting.
            ExportFontsAsBase64 = true,          // Embed fonts in Base64 inside CSS.
            CssStyleSheetType = CssStyleSheetType.Embedded,
            PrettyFormat = true
        };

        // Save the document as HTML using the configured options.
        string htmlPath = Path.Combine(outputFolder, "Sample.html");
        doc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // Validation: ensure the HTML file exists and contains Base64 data.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("base64"))
            throw new InvalidOperationException("The HTML does not contain Base64-encoded font data.");

        // Indicate successful completion.
        Console.WriteLine("Document successfully saved as HTML with embedded Base64 fonts at:");
        Console.WriteLine(htmlPath);
    }
}
