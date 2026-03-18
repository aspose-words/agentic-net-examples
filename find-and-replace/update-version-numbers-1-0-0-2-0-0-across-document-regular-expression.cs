using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class UpdateVersionNumbers
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Define a regular expression that matches the exact version string "1.0.0".
        Regex versionPattern = new Regex(@"\b1\.0\.0\b");

        // Replace all occurrences of the pattern with the new version "2.0.0".
        // Using the overload that accepts a Regex and a replacement string.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementsMade = doc.Range.Replace(versionPattern, "2.0.0", options);

        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
