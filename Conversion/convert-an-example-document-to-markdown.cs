using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToMarkdown
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words).
        string inputPath = @"C:\Docs\Example.docx";

        // Path where the Markdown file will be saved.
        string outputPath = @"C:\Docs\Example.md";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create Markdown save options.
        // Here we demonstrate exporting tables as raw HTML within the Markdown.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the document as Markdown using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
