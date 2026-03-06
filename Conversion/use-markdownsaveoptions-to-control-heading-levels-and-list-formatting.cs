using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Ensure input and output directories exist.
        Directory.CreateDirectory(MyDir);
        Directory.CreateDirectory(ArtifactsDir);

        // Load the source DOCX document.
        Document doc = new Document(Path.Combine(MyDir, "Sample.docx"));

        // Create Markdown save options.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions();

        // Control list formatting: export lists using Markdown syntax.
        mdOptions.ListExportMode = MarkdownListExportMode.MarkdownSyntax;

        // Example: export tables that cannot be represented in pure Markdown as raw HTML.
        mdOptions.ExportAsHtml = MarkdownExportAsHtml.NonCompatibleTables;

        // Save the document as a Markdown file using the configured options.
        doc.Save(Path.Combine(ArtifactsDir, "Converted.md"), mdOptions);
    }

    // Paths to the input and output folders (adjust as needed).
    private static readonly string MyDir = Path.Combine(Directory.GetCurrentDirectory(), "Input");
    private static readonly string ArtifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
}
