using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class LongWordReplaceEvaluator : IReplacingCallback
{
    // Replace only words whose length is greater than ten characters.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // If the matched word is longer than 10 characters, replace it.
        if (args.Match.Value.Length > 10)
        {
            // Example replacement text – you can change this as needed.
            args.Replacement = "REPLACED";
            return ReplaceAction.Replace;
        }

        // Otherwise skip the replacement for this match.
        return ReplaceAction.Skip;
    }
}

class Program
{
    static void Main()
    {
        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This sentence contains supercalifragilisticexpialidocious and short words.");
        builder.Writeln("Anotherexampleword demonstrates the replacement.");

        // Set up find/replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new LongWordReplaceEvaluator());

        // Use a regular expression that matches any whole word.
        Regex wordRegex = new Regex(@"\b\w+\b");

        // Perform the replace operation. The callback decides which matches to replace.
        int replacedCount = doc.Range.Replace(wordRegex, string.Empty, options);

        // Output the number of replacements performed.
        Console.WriteLine($"Replacements made: {replacedCount}");

        // Save the modified document.
        doc.Save("LongWordReplaced.docx");
    }
}
