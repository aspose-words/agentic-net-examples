using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing document (replace with your actual file path)
            Document doc = new Document("Input.docx");

            // The phrase we want to search for
            string phrase = "specific phrase";

            // Check if any paragraph in the document contains the phrase
            bool containsPhrase = doc.FirstSection.Body.Paragraphs
                .Any(paragraph => paragraph.GetText().IndexOf(phrase, StringComparison.OrdinalIgnoreCase) >= 0);

            Console.WriteLine(containsPhrase
                ? $"The document contains the phrase \"{phrase}\"."
                : $"The phrase \"{phrase}\" was not found in the document.");

            // Save the document (optional, using the provided save rule)
            doc.Save("Output.docx");
        }
    }
}
