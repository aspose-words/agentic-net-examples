using System;
using System.IO;
using Aspose.Words;

namespace ExtractParagraphContent
{
    class Program
    {
        static void Main()
        {
            // Path to the DOT (Word template) file.
            string inputPath = @"C:\Docs\Template.dot";

            // Load the DOT file into an Aspose.Words Document.
            Document doc = new Document(inputPath);

            // Get the collection of paragraphs from the main body of the first section.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Iterate through the paragraphs and extract the text between them.
            // Here we simply output each paragraph's text (excluding the trailing paragraph break).
            for (int i = 0; i < paragraphs.Count; i++)
            {
                // Get the full text of the paragraph, which includes the paragraph break character.
                string paragraphTextWithBreak = paragraphs[i].GetText();

                // Trim the trailing paragraph break characters (\r, \f, or \a) for clean output.
                string cleanText = paragraphTextWithBreak.TrimEnd('\r', '\f', '\a');

                Console.WriteLine($"Paragraph {i + 1}: {cleanText}");
            }

            // Optionally, write the extracted content to a plain text file.
            string outputPath = @"C:\Docs\ExtractedContent.txt";
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    string cleanText = paragraphs[i].GetText().TrimEnd('\r', '\f', '\a');
                    writer.WriteLine(cleanText);
                }
            }

            Console.WriteLine("Extraction complete.");
        }
    }
}
