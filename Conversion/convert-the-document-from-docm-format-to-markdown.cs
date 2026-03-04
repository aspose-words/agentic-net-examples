using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCM file path
        string inputPath = "input.docm";

        // Output Markdown file path
        string outputPath = "output.md";

        // Load the DOCM document
        Document doc = new Document(inputPath);

        // Configure Markdown save options
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        // Explicitly set the target format (optional, as the extension also determines it)
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as Markdown
        doc.Save(outputPath, saveOptions);
    }
}
