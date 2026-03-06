using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithRegex
{
    static void Main()
    {
        // Path to the input DOCX file.
        const string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the modified DOCX will be saved.
        const string outputPath = @"C:\Docs\OutputDocument.docx";

        // Regular expression pattern to search for.
        const string regexPattern = @"\b[Aa]spose\b";

        // Replacement text.
        const string replacement = "Aspose.Words";

        // Set to true for case‑sensitive search, false for case‑insensitive.
        bool matchCase = true;

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = matchCase          // Toggle case sensitivity.
        };

        // Perform the regex replace.
        int replacementsMade = doc.Range.Replace(new Regex(regexPattern), replacement, options);

        // Optional: display how many replacements were performed.
        Console.WriteLine($"{replacementsMade} replacement(s) made.");

        // Save the updated document.
        doc.Save(outputPath);
    }
}
