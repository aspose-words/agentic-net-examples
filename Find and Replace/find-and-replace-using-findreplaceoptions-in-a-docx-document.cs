using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceDemo
{
    static void Main()
    {
        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Ruby bought a ruby necklace.");
        builder.Writeln("Jackson will meet you in Jacksonville.");
        builder.Writeln("Numbers 1, 2, 3.");

        // -----------------------------------------------------------------
        // Example 1: Simple case‑sensitive replace using FindReplaceOptions.
        // -----------------------------------------------------------------
        FindReplaceOptions optionsCase = new FindReplaceOptions
        {
            MatchCase = true               // Enable case‑sensitive matching.
        };
        // Replace only the capitalized "Ruby".
        doc.Range.Replace("Ruby", "Jade", optionsCase);

        // -----------------------------------------------------------------
        // Example 2: Whole‑word only replace.
        // -----------------------------------------------------------------
        FindReplaceOptions optionsWholeWord = new FindReplaceOptions
        {
            FindWholeWordsOnly = true      // Replace only whole words.
        };
        // Replace "Jackson" but not the "Jacksonville" part.
        doc.Range.Replace("Jackson", "Louis", optionsWholeWord);

        // -----------------------------------------------------------------
        // Example 3: Regular‑expression replace with substitutions.
        // -----------------------------------------------------------------
        FindReplaceOptions optionsRegex = new FindReplaceOptions
        {
            UseSubstitutions = true,       // Enable $1, $2 … substitutions.
            LegacyMode = false             // Required for substitution support.
        };
        Regex regex = new Regex(@"Numbers (\d+), (\d+), (\d+)");
        // Rearrange the numbers.
        doc.Range.Replace(regex, @"$3, $2, $1", optionsRegex);

        // Save the modified document.
        doc.Save("FindReplaceResult.docx");
    }
}
