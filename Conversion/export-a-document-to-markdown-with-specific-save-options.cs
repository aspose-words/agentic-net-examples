using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MarkdownExportExample
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Configure the Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Export tables as raw HTML.
            ExportAsHtml = MarkdownExportAsHtml.Tables,
            // Export list items using Markdown syntax.
            ListExportMode = MarkdownListExportMode.MarkdownSyntax,
            // Export links as reference style.
            LinkExportMode = MarkdownLinkExportMode.Reference,
            // Export underline formatting as "++".
            ExportUnderlineFormatting = true,
            // Export OfficeMath objects as LaTeX.
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex,
            // Explicitly set the format to Markdown.
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document to a Markdown file using the configured options.
        doc.Save("Output.md", saveOptions);
    }
}
