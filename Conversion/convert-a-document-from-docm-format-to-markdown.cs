using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocmToMarkdown
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCM file path
            string inputPath = @"C:\Docs\Sample.docm";

            // Output Markdown file path
            string outputPath = @"C:\Docs\Sample.md";

            // Load the DOCM document. The constructor automatically detects the format.
            Document doc = new Document(inputPath);

            // Save the document as Markdown using the SaveFormat enumeration.
            // This uses the built‑in save method; no custom save logic is required.
            doc.Save(outputPath, SaveFormat.Markdown);

            // Alternatively, you can use MarkdownSaveOptions for more control:
            // var saveOptions = new MarkdownSaveOptions();
            // doc.Save(outputPath, saveOptions);
        }
    }
}
