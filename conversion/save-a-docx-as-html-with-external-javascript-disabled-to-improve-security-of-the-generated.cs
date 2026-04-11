using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary DOCX and the resulting HTML.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.docx");
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.html");
        string cssPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.css");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some regular text.
        builder.Writeln("This is a sample document.");

        // Add a hyperlink that contains JavaScript (potential XSS vector).
        builder.InsertHyperlink("Click me", "javascript:alert('XSS')", false);
        builder.Writeln();

        // Save the DOCX to disk (required by the rule to have a local source file).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX back (simulating an existing file scenario).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure HTML save options to disable JavaScript in links.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Replace JavaScript links with a harmless placeholder.
            RemoveJavaScriptFromLinks = true,

            // Store CSS in an external file for clarity.
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = cssPath
        };

        // Save the document as HTML using the configured options.
        loadedDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Validation: ensure the HTML file was created and contains no
        //    active JavaScript code (the placeholder "javascript:void(0)" is allowed).
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);

        // Look for any JavaScript link that is not the safe placeholder.
        if (htmlContent.IndexOf("javascript:alert", StringComparison.OrdinalIgnoreCase) >= 0 ||
            htmlContent.IndexOf("javascript:", StringComparison.OrdinalIgnoreCase) >= 0 &&
            !htmlContent.Contains("javascript:void(0)", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("JavaScript links were not properly removed from the HTML output.");
        }

        // Optional cleanup (commented out to allow inspection of output files).
        // File.Delete(docxPath);
        // File.Delete(htmlPath);
        // File.Delete(cssPath);
    }
}
