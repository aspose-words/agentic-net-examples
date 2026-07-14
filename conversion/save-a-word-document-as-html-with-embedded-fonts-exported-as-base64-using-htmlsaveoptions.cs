using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text and set a specific font.
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document with embedded fonts.");

        // Configure HtmlSaveOptions to embed fonts as Base64.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontResources = true,      // Export font resources.
            ExportFontsAsBase64 = true,      // Embed fonts in Base64 within the CSS.
            CssStyleSheetType = CssStyleSheetType.Embedded,
            PrettyFormat = true
        };

        // Define the output HTML file path.
        string outputPath = "output.html";

        // Save the document as HTML using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Verify that the HTML contains Base64‑encoded font data.
        string htmlContent = File.ReadAllText(outputPath);
        // Look for a data URI that represents an embedded font (case‑insensitive).
        bool containsBase64Font = htmlContent.Contains("data:font", StringComparison.OrdinalIgnoreCase) ||
                                  htmlContent.Contains("data:application", StringComparison.OrdinalIgnoreCase);
        if (!containsBase64Font)
            throw new InvalidOperationException("The HTML does not contain embedded Base64 font data.");
    }
}
