using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOCX file path
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Output Markdown file path
            string outputPath = @"C:\Docs\SampleDocument.md";

            // Load the DOCX document using the Document constructor (lifecycle rule)
            Document doc = new Document(inputPath);

            // Create MarkdownSaveOptions to specify Markdown format (lifecycle rule)
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the save format to Markdown
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document as Markdown using the Save method with options (lifecycle rule)
            doc.Save(outputPath, saveOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
