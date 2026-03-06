using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains Structured Document Tags (SDTs).
        Document doc = new Document("StructuredDocumentTags.docx");

        // Define the text to search for and its replacement.
        // This example replaces the literal string "Placeholder" with "New Text".
        string pattern = "Placeholder";
        string replacement = "New Text";

        // Configure find/replace options.
        // Setting IgnoreStructuredDocumentTags to true treats the content of each SDT as plain text,
        // allowing the pattern to be found even when it spans across SDT boundaries.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreStructuredDocumentTags = true
        };

        // Execute the replace operation on the entire document range.
        int replacementsMade = doc.Range.Replace(pattern, replacement, options);
        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save("StructuredDocumentTags_Replaced.docx");
    }
}
