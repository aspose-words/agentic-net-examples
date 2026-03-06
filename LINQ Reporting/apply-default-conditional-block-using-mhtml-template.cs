using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class MhtmlConditionalExample
{
    static void Main()
    {
        // Paths to input HTML, template and output MHTML files.
        string inputHtmlPath = @"C:\Docs\InputConditional.html";
        string templatePath = @"C:\Docs\DefaultTemplate.dotx";
        string outputMhtmlPath = @"C:\Docs\Result.mht";

        // -----------------------------------------------------------------
        // Load the HTML document with options that enable processing of
        // conditional comments (e.g., VML conditional blocks).
        // -----------------------------------------------------------------
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.SupportVml = true;               // Enable VML conditional handling.
        loadOptions.BlockImportMode = BlockImportMode.Preserve; // Preserve block properties.

        Document doc = new Document(inputHtmlPath, loadOptions);

        // -----------------------------------------------------------------
        // Ensure the document uses a default template when saving.
        // This is required because the document does not have an attached template.
        // -----------------------------------------------------------------
        doc.AutomaticallyUpdateStyles = true;        // Allow style updates from a template.
        doc.AttachedTemplate = string.Empty;         // No explicit template attached.

        // -----------------------------------------------------------------
        // Prepare save options for MHTML output.
        // ExportCidUrlsForMhtmlResources is set to true to improve compatibility
        // with mail agents that expect CID URLs.
        // DefaultTemplate points to the template that will be applied during save.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true,
            DefaultTemplate = templatePath,
            CssStyleSheetType = CssStyleSheetType.External,
            ExportFontResources = true,
            PrettyFormat = true
        };

        // Save the document as MHTML using the configured options.
        doc.Save(outputMhtmlPath, saveOptions);
    }
}
