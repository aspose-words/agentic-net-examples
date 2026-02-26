using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // Example 1: Simple string replace with case‑sensitive, whole‑word only.
        // -----------------------------------------------------------------
        FindReplaceOptions stringOptions = new FindReplaceOptions
        {
            MatchCase = true,               // Enable case‑sensitive matching.
            FindWholeWordsOnly = true,      // Replace only whole words.
            IgnoreFootnotes = false         // Include footnotes in the search.
        };

        // Replace every occurrence of "Apple" with "Orange".
        int stringReplacements = doc.Range.Replace("Apple", "Orange", stringOptions);
        Console.WriteLine($"String replacements performed: {stringReplacements}");

        // -----------------------------------------------------------------
        // Example 2: Regular‑expression replace with captured groups.
        // -----------------------------------------------------------------
        FindReplaceOptions regexOptions = new FindReplaceOptions
        {
            UseSubstitutions = true,   // Enable $1, $2 … substitutions.
            LegacyMode = false         // Required for advanced features.
        };

        // Swap the order of two words (e.g., "Hello World" -> "World Hello").
        Regex regex = new Regex(@"(\w+)\s+(\w+)");
        int regexReplacements = doc.Range.Replace(regex, "$2 $1", regexOptions);
        Console.WriteLine($"Regex replacements performed: {regexReplacements}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
