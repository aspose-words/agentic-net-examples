using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocmToMarkdown
{
    static void Main()
    {
        // Load the DOCM file.
        Document doc = new Document("Input.docm");

        // Create Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Export tables that cannot be represented in pure Markdown as raw HTML.
        saveOptions.ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables;

        // Export list items as plain text.
        saveOptions.ListExportMode = MarkdownListExportMode.PlainText;

        // Set image resolution to 300 DPI.
        saveOptions.ImageResolution = 300;

        // Explicitly set the save format to Markdown (optional, default for this options class).
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as a Markdown file using the configured options.
        doc.Save("Output.md", saveOptions);
    }
}
