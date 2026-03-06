using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToMarkdownConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "Document.docx";

        // Path where the resulting Markdown file will be saved.
        string outputPath = "Document.md";

        // Load the DOCX document.
        Document doc = new Document(sourcePath);

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Explicitly set the format to Markdown (optional, but clarifies intent).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown.
        doc.Save(outputPath, saveOptions);
    }
}
