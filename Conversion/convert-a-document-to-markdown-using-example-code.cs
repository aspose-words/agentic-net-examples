using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (DOCX, DOC, etc.).
        Document doc = new Document("InputDocument.docx");

        // Create Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Specify that the output format is Markdown.
        // This is optional because MarkdownSaveOptions already implies Markdown,
        // but setting it explicitly follows the example pattern.
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as a Markdown file.
        doc.Save("OutputDocument.md", saveOptions);
    }
}
