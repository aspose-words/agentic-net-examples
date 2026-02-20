using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("Input.docm");

        // Create Markdown save options and customize them as needed.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Example: export tables that cannot be represented in pure Markdown as raw HTML.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,

            // Example: export links as reference blocks.
            LinkExportMode = MarkdownLinkExportMode.Reference,

            // Example: preserve empty paragraphs as empty lines.
            EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

            // Example: embed generator name in the output.
            ExportGeneratorName = true
        };

        // Save the document as Markdown using the customized options.
        doc.Save("Output.md", saveOptions);
    }
}
