using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputPath = "input.docx";

        // Path where the Markdown file will be saved
        string outputPath = "output.md";

        // Load the document from the file system
        Document doc = new Document(inputPath);

        // Create save options for Markdown format
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Optional: customize options, e.g. export images as Base64
        // saveOptions.ExportImagesAsBase64 = true;
        // saveOptions.ListExportMode = MarkdownListExportMode.MarkdownSyntax;

        // Save the document as a Markdown file using the specified options
        doc.Save(outputPath, saveOptions);
    }
}
