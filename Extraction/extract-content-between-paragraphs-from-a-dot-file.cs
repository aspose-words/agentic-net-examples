using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;

namespace ExtractParagraphContent
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the DOT (Word template) file.
            string dotFilePath = @"C:\Docs\Template.dot";

            // Load the DOT file into an Aspose.Words Document.
            Document doc = new Document(dotFilePath);

            // Retrieve the collection of paragraphs from the main body of the first section.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // List to hold the text of each paragraph (without the trailing paragraph break).
            List<string> paragraphTexts = new List<string>();

            // Iterate through all paragraphs and extract their visible text.
            foreach (Paragraph para in paragraphs)
            {
                // GetText() returns the paragraph text including the end-of-paragraph character.
                // TrimEnd removes the trailing control characters (\r, \f, etc.).
                string text = para.GetText().TrimEnd('\r', '\n', '\f', '\v');
                paragraphTexts.Add(text);
            }

            // Example: concatenate the texts of consecutive paragraphs, inserting a custom separator.
            // This demonstrates extracting "content between paragraphs".
            StringBuilder resultBuilder = new StringBuilder();
            for (int i = 0; i < paragraphTexts.Count - 1; i++)
            {
                // Content between paragraph i and paragraph i+1 is the text of paragraph i+1.
                // Here we simply join them with a line separator for demonstration.
                resultBuilder.AppendLine(paragraphTexts[i]);
                resultBuilder.AppendLine("--- Between ---");
                resultBuilder.AppendLine(paragraphTexts[i + 1]);
                resultBuilder.AppendLine();
            }

            // Output the extracted content to the console.
            Console.WriteLine("Extracted content between paragraphs:");
            Console.WriteLine(resultBuilder.ToString());

            // Optionally, save the extracted content to a plain text file.
            string outputPath = @"C:\Docs\ExtractedContent.txt";
            System.IO.File.WriteAllText(outputPath, resultBuilder.ToString());
        }
    }
}
