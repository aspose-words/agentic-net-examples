using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Create Markdown save options. This object controls how the document is exported to Markdown.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Explicitly set the save format to Markdown (optional, but ensures correctness).
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as a Markdown file using the specified options.
        doc.Save("output.md", saveOptions);
    }
}
