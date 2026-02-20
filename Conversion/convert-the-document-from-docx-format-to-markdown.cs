using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocxToMarkdown
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCX file path
            string inputPath = @"C:\Docs\input.docx";

            // Output Markdown file path
            string outputPath = @"C:\Docs\output.md";

            // Load the DOCX document
            Document doc = new Document(inputPath);

            // Create Markdown save options (default settings)
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

            // Save the document as Markdown
            doc.Save(outputPath, saveOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
