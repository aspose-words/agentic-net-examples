using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path to the destination Markdown file.
        string outputPath = "output.md";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Create Markdown save options (customize if needed).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Save the document as Markdown.
        doc.Save(outputPath, saveOptions);
    }
}
