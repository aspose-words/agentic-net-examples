using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the input Markdown file.
        string inputPath = @"C:\Docs\Input.md";

        // Load the Markdown document into an Aspose.Words Document.
        // Use MarkdownLoadOptions to preserve empty lines if needed.
        var loadOptions = new MarkdownLoadOptions
        {
            PreserveEmptyLines = true
        };
        Document doc = new Document(inputPath, loadOptions);

        // Create a DocumentBuilder to manipulate the document if required.
        // For example, insert a heading before the existing content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentStart();
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Generated Markdown Document");

        // Configure Markdown save options.
        var saveOptions = new MarkdownSaveOptions
        {
            // Export tables as raw HTML to preserve complex structures.
            ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables,
            // Preserve empty paragraphs as empty lines.
            EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,
            // Use pretty formatting for readability.
            PrettyFormat = true,
            // Export links as reference style.
            LinkExportMode = MarkdownLinkExportMode.Reference,
            // Align table contents to the center.
            TableContentAlignment = TableContentAlignment.Center
        };

        // Path to the output Markdown file.
        string outputPath = @"C:\Docs\Output.md";

        // Save the document as Markdown using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
