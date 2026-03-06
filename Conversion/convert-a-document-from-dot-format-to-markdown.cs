using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = "Template.dot";

        // Path where the resulting Markdown file will be saved.
        string outputPath = "Output.md";

        // Load the DOT document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Configure save options for Markdown output.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Explicitly set the format to Markdown (optional, as the class defaults to this format).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as a Markdown file using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
