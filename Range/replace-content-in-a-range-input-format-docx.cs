using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceInRange
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the modified DOCX will be saved.
        string outputPath = "output.docx";

        // Load the existing document (lifecycle: load rule).
        Document doc = new Document(inputPath);

        // Define the text to find and its replacement.
        string pattern = "OldText";
        string replacement = "NewText";

        // Optional: configure find/replace options (e.g., case‑insensitive, whole word).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replace operation on the document's main range (range replace rule).
        int count = doc.Range.Replace(pattern, replacement, options);

        // Optionally, report how many replacements were made.
        Console.WriteLine($"{count} occurrence(s) replaced.");

        // Save the modified document (lifecycle: save rule).
        doc.Save(outputPath);
    }
}
