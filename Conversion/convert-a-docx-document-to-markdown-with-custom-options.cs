using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocxToMarkdown
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the resulting Markdown file will be saved.
        string outputPath = @"C:\Docs\OutputDocument.md";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure the Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Embed images directly in the Markdown file as Base64 data URIs.
            ExportImagesAsBase64 = true,

            // Export tables that cannot be represented in pure Markdown as raw HTML.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

            // Export OfficeMath objects as LaTeX markup.
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,

            // Export all links using reference style.
            LinkExportMode = MarkdownLinkExportMode.Reference,

            // Set the image resolution (relevant when images are saved as separate files).
            ImageResolution = 300,

            // Explicitly set the save format to Markdown.
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as a Markdown file using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
