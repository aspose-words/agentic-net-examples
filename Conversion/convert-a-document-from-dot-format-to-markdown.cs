using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class DotToMarkdownConverter
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = @"C:\Docs\Template.dot";

        // Path where the resulting Markdown file will be saved.
        string outputPath = @"C:\Docs\Converted.md";

        // Load the DOT file. Explicitly specify the format to avoid auto‑detection.
        var loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Dot
        };
        Document doc = new Document(inputPath, loadOptions);

        // Configure Markdown save options if needed. The class does not expose a
        // PreserveEmptyLines property, so we use the default options or set other
        // supported properties here.
        var saveOptions = new MarkdownSaveOptions();

        // Save the document as Markdown.
        doc.Save(outputPath, saveOptions);
    }
}
