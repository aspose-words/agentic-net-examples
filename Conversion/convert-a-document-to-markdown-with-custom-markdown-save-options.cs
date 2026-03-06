using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToMarkdown
{
    static void Main()
    {
        // Paths to the source document and the resulting Markdown file.
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Docs\Output.md";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create a MarkdownSaveOptions instance via the factory method.
        SaveOptions genericOptions = SaveOptions.CreateSaveOptions(SaveFormat.Markdown);
        MarkdownSaveOptions mdOptions = (MarkdownSaveOptions)genericOptions;

        // Customize the Markdown conversion.
        mdOptions.ExportAsHtml = MarkdownExportAsHtml.Tables;               // Export tables as raw HTML.
        mdOptions.ImageResolution = 300;                                    // Use higher DPI for exported images.
        mdOptions.ListExportMode = MarkdownListExportMode.PlainText;        // Export lists as plain text.
        mdOptions.EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine; // Preserve empty paragraphs.
        mdOptions.ExportGeneratorName = false;                              // Omit Aspose.Words generator comment.

        // Save the document as Markdown using the customized options.
        doc.Save(outputPath, mdOptions);
    }
}
