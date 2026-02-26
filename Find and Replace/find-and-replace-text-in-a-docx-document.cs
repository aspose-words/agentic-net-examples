using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Simple string replacement (case‑insensitive, whole document).
        // Replaces all occurrences of the placeholder "_FullName_" with "John Doe".
        int count = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Replacements made: {count}");

        // Example of using FindReplaceOptions for more control:
        // Replace every period that is immediately before a paragraph break with an exclamation point,
        // and right‑align the paragraphs that contain the match.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;
        int countWithOptions = doc.Range.Replace(".&p", "!&p", options);
        Console.WriteLine($"Replacements with options made: {countWithOptions}");

        // Example of a regular‑expression replacement:
        // Replace all numbers with a paragraph break.
        int regexCount = doc.Range.Replace(new Regex(@"\d+"), "&p");
        Console.WriteLine($"Regex replacements made: {regexCount}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
