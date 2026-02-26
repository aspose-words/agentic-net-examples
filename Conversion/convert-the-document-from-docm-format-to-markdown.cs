using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = "Input.docm";

        // Path where the Markdown file will be saved.
        string outputPath = "Output.md";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Set up Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        saveOptions.SaveFormat = SaveFormat.Markdown; // Explicitly specify Markdown format.

        // Save the document as a Markdown file.
        doc.Save(outputPath, saveOptions);
    }
}
