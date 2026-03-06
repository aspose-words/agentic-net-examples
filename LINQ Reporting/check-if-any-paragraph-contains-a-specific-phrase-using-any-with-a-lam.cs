using System;
using System.Linq;
using Aspose.Words;

class CheckParagraphPhrase
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // The phrase we are looking for.
        const string phrase = "specific phrase";

        // Use LINQ Any with a lambda to determine if any paragraph contains the phrase.
        bool phraseFound = doc.FirstSection.Body.Paragraphs
            .Any(paragraph => paragraph.GetText().Contains(phrase, StringComparison.OrdinalIgnoreCase));

        // Output the result.
        Console.WriteLine(phraseFound
            ? $"The phrase \"{phrase}\" was found in at least one paragraph."
            : $"The phrase \"{phrase}\" was not found in any paragraph.");

        // Optionally save the document (no changes made, just demonstrating the save rule).
        doc.Save("Result.docx");
    }
}
