using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class DocmToMarkdownConverter
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourcePath = "input.docm";

        // Path where the resulting Markdown file will be saved.
        string destinationPath = "output.md";

        // Load the DOCM document. Explicitly set the LoadFormat to Docm to ensure correct detection.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docm
        };
        Document document = new Document(sourcePath, loadOptions);

        // Create Markdown save options. Default settings are sufficient for a basic conversion.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Save the document as Markdown.
        document.Save(destinationPath, saveOptions);
    }
}
