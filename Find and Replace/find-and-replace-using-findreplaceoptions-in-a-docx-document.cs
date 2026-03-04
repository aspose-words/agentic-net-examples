using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceDemo
{
    static void Main()
    {
        // Create a new document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.Writeln("Hello world again.");
        builder.Writeln("HELLO WORLD!");

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Perform a case‑sensitive search.
            MatchCase = true,
            // Replace only whole word matches.
            FindWholeWordsOnly = true
        };

        // Replace the word "Hello" with "Greetings" using the options.
        int replacements = doc.Range.Replace("Hello", "Greetings", options);
        Console.WriteLine($"Replacements made: {replacements}");

        // Save the modified document.
        doc.Save("FindReplaceResult.docx");
    }
}
