using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample markdown-like content.
        builder.Writeln("## Sample Markdown Export");
        builder.Writeln("This is a paragraph with **bold** text.");
        builder.Writeln();

        // Insert a simple table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Export tables as raw HTML to preserve complex structures.
            ExportAsHtml = MarkdownExportAsHtml.Tables,

            // Export lists using standard markdown syntax.
            ListExportMode = MarkdownListExportMode.MarkdownSyntax,

            // Export links as reference style.
            LinkExportMode = MarkdownLinkExportMode.Reference,

            // Preserve empty paragraphs as empty lines.
            EmptyParagraphExportMode = MarkdownEmptyParagraphExportMode.EmptyLine,

            // Use UTF-8 encoding without BOM.
            Encoding = new UTF8Encoding(false),

            // Do not embed the Aspose.Words generator name.
            ExportGeneratorName = false,

            // Save images as Base64 strings.
            ExportImagesAsBase64 = true,

            // If images were saved as files, specify folder and alias.
            ImagesFolder = "Images",
            ImagesFolderAlias = "images"
        };

        // Save the document to a Markdown file using the configured options.
        doc.Save("Output.md", saveOptions);
    }
}
