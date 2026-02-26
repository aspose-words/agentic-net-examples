using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindAndReplaceDemo
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Simple string replace (case‑insensitive, whole document).
        // Replaces all occurrences of the placeholder "_FullName_" with "John Doe".
        int count = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Simple replace made {count} replacements.");

        // Example of using FindReplaceOptions to control the replace operation.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Perform a case‑sensitive replace.
            MatchCase = true,
            // Replace only whole words.
            FindWholeWordsOnly = true,
            // Apply a paragraph format to paragraphs that contain a match.
            ApplyParagraphFormat = { Alignment = ParagraphAlignment.Right }
        };

        // Replace every occurrence of "Important" with "Critical" using the options above.
        int advancedCount = doc.Range.Replace("Important", "Critical", options);
        Console.WriteLine($"Advanced replace made {advancedCount} replacements.");

        // Example of a regular‑expression replace.
        // Replace any sequence of digits with a paragraph break.
        int regexCount = doc.Range.Replace(new Regex(@"\d+"), "&p", options);
        Console.WriteLine($"Regex replace made {regexCount} replacements.");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
