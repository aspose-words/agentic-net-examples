using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class MarkdownConverter
{
    static void Main()
    {
        // Path to the source Markdown file.
        string markdownPath = @"C:\Docs\source.md";

        // Path to the output HTML file.
        string htmlPath = @"C:\Docs\converted.html";

        // Load the Markdown document with custom load options.
        var loadOptions = new MarkdownLoadOptions
        {
            // Preserve empty lines in the source Markdown.
            PreserveEmptyLines = true,

            // Resolve relative URIs based on this base URI (optional).
            BaseUri = "https://example.com/resources/"
        };

        // Create a Document object from the Markdown file.
        Document doc = new Document(markdownPath, loadOptions);

        // (Optional) Manipulate the document here, e.g., modify content controls.
        // For demonstration, we will just ensure the document is not empty.
        if (doc.GetText().Trim().Length == 0)
        {
            Console.WriteLine("The loaded Markdown document is empty.");
            return;
        }

        // Configure save options for HTML output.
        var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export the document generator name (Aspose.Words) into the output.
            ExportGeneratorName = true,

            // Export images as Base64 to keep a single HTML file.
            ExportImagesAsBase64 = true,

            // Use pretty formatting for readability.
            PrettyFormat = true
        };

        // Save the document as HTML.
        doc.Save(htmlPath, saveOptions);

        Console.WriteLine($"Markdown file '{markdownPath}' was successfully converted to HTML at '{htmlPath}'.");
    }
}
