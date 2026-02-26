using System;
using Aspose.Words;

namespace ParagraphLineCounter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input DOCX file.
            string inputPath = "InputDocument.docx";

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Update word count and also calculate the number of lines.
            doc.UpdateWordCount(true);

            // Retrieve the built‑in properties for lines and paragraphs.
            int lineCount = doc.BuiltInDocumentProperties.Lines;
            int paragraphCount = doc.BuiltInDocumentProperties.Paragraphs;

            // Output the results.
            Console.WriteLine($"Lines in document: {lineCount}");
            Console.WriteLine($"Paragraphs in document: {paragraphCount}");

            // Optionally, save the document (e.g., after updating properties).
            string outputPath = "OutputDocument.docx";
            doc.Save(outputPath);
        }
    }
}
