using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToMarkdownConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "input.doc";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the DOC document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Create save options for Markdown. Default options are sufficient for a basic conversion.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Save the document as Markdown using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
