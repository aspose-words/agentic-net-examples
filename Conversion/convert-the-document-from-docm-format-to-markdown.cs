using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class DocmToMarkdownConverter
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\source.docm";

        // Path where the resulting Markdown file will be saved.
        string outputPath = @"C:\Docs\converted.md";

        // Load the DOCM document. The LoadFormat is automatically detected,
        // but we can explicitly set it to Docm for clarity.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docm
        };
        Document doc = new Document(inputPath, loadOptions);

        // Configure Markdown save options if needed (defaults are sufficient for a basic conversion).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Example: preserve empty lines as empty lines in the output.
            EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,
            // Example: export lists using Markdown syntax.
            ListExportMode = MarkdownListExportMode.MarkdownSyntax
        };

        // Save the document as Markdown.
        doc.Save(outputPath, saveOptions);
    }
}
