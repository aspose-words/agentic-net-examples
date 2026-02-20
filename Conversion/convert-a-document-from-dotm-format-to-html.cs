using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DotmToHtmlConverter
{
    static void Main()
    {
        // Path to the source DOTM file.
        string inputPath = @"C:\Docs\Template.dotm";

        // Path where the resulting HTML file will be saved.
        string outputPath = @"C:\Docs\Template.html";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Configure HTML save options (optional – can be omitted for defaults).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Example: embed images as Base64 to keep a single HTML file.
            ExportImagesAsBase64 = true,
            // Example: embed CSS directly into the HTML.
            CssSavingCallback = null,
            // Example: use XHTML 1.0 Transitional output.
            HtmlVersion = HtmlVersion.Xhtml,
            ExportXhtmlTransitional = true,
            PrettyFormat = true
        };

        // Save the document as HTML.
        doc.Save(outputPath, saveOptions);
    }
}
