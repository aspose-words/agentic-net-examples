using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (any format supported by Aspose.Words).
            const string inputFile = @"C:\Docs\SampleDocument.docx";

            // Path where the Markdown file will be saved.
            const string outputFile = @"C:\Docs\SampleDocument.md";

            // Load the document using the standard Document constructor.
            Document doc = new Document(inputFile);

            // Create MarkdownSaveOptions and specify that the output format is Markdown.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown
            };

            // Save the document to Markdown using the Save method that accepts SaveOptions.
            doc.Save(outputFile, saveOptions);
        }
    }
}
