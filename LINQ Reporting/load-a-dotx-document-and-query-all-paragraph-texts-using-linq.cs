using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTX template file.
            // The Document(string) constructor is the prescribed load rule.
            Document doc = new Document(@"C:\Templates\SampleTemplate.dotx");

            // Retrieve all Paragraph nodes in the document (including those in headers/footers).
            // GetChildNodes(NodeType.Paragraph, true) returns a live collection of all paragraphs.
            // LINQ is used to project each Paragraph to its text content.
            List<string> paragraphTexts = doc.GetChildNodes(NodeType.Paragraph, true)
                                            .Cast<Paragraph>()
                                            .Select(p => p.GetText().Trim())
                                            .ToList();

            // Output the collected paragraph texts.
            foreach (string text in paragraphTexts)
            {
                Console.WriteLine(text);
            }
        }
    }
}
