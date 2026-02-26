using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertListToPlainTextMarkdown
{
    static void Main()
    {
        // Path to the source DOC document that contains the list.
        string inputPath = @"C:\Docs\ListDocument.doc";

        // Path where the resulting Markdown file will be saved.
        string outputPath = @"C:\Docs\ListDocument.md";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Configure Markdown save options to export list items as plain text.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ListExportMode = MarkdownListExportMode.PlainText
        };

        // Save the document as Markdown using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
