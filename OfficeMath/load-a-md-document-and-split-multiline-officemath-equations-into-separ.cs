using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class SplitMultilineOfficeMath
{
    static void Main()
    {
        // Path to the source Markdown file that may contain multiline OfficeMath equations.
        string inputPath = @"C:\Docs\source.md";

        // Load the Markdown document while preserving empty lines.
        // This ensures that any line breaks inside the document are retained.
        MarkdownLoadOptions loadOptions = new MarkdownLoadOptions
        {
            PreserveEmptyLines = true
        };
        Document doc = new Document(inputPath, loadOptions);

        // Configure save options to export OfficeMath as MarkItDown (LaTeX compatible).
        // This format writes each OfficeMath equation on its own line in the resulting Markdown.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            OfficeMathExportMode = MarkdownOfficeMathExportMode.MarkItDown
        };

        // Save the processed document. Each multiline OfficeMath equation will be split
        // into separate lines according to the MarkItDown export mode.
        string outputPath = @"C:\Docs\output.md";
        doc.Save(outputPath, saveOptions);
    }
}
