using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputPath = "input.docx";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Set up Markdown save options to exclude headers/footers.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None,
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as a Markdown file using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
