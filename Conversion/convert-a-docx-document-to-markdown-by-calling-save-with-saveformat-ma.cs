using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Path where the resulting Markdown file will be saved.
            string outputPath = @"C:\Docs\SampleDocument.md";

            // Load the DOCX document using the Document constructor.
            Document doc = new Document(inputPath);

            // Save the document in Markdown format.
            doc.Save(outputPath, SaveFormat.Markdown);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
