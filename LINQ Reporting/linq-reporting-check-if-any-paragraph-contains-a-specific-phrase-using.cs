using System;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsLinqExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTX template document.
            Document doc = new Document("Template.dotx");

            // The phrase we want to search for.
            string phrase = "specific phrase";

            // Use LINQ Any with a lambda to determine if any paragraph contains the phrase.
            bool containsPhrase = doc.FirstSection.Body.Paragraphs
                .Any(paragraph => paragraph.GetText().IndexOf(phrase, StringComparison.OrdinalIgnoreCase) >= 0);

            // Output the result.
            Console.WriteLine(containsPhrase
                ? $"The document contains the phrase \"{phrase}\"."
                : $"The phrase \"{phrase}\" was not found in the document.");

            // Optionally save the document (no changes made in this example).
            doc.Save("Result.docx");
        }
    }
}
