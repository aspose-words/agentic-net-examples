using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("Input.docm");

        // Create and configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        // Export tables that cannot be represented in pure Markdown as raw HTML.
        saveOptions.ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables;
        // Set a higher image resolution for exported images.
        saveOptions.ImageResolution = 300;
        // Explicitly set the format to Markdown (optional, default when using this options class).
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as a Markdown file using the configured options.
        doc.Save("Output.md", saveOptions);
    }
}
