using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document that contains SmartTags.
        Document doc = new Document("InputWithSmartTags.docx");

        // Define the text to find and its replacement.
        string pattern = "_FullName_";          // Example placeholder inside a SmartTag.
        string replacement = "John Doe";

        // Configure FindReplaceOptions to treat the content of StructuredDocumentTag
        // (including SmartTag) as simple text, allowing the replace operation to work
        // across SmartTag boundaries.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreStructuredDocumentTags = true   // Important for SmartTag replacement.
        };

        // Perform the find-and-replace operation on the whole document range.
        int replacedCount = doc.Range.Replace(pattern, replacement, options);

        // Optional: output how many replacements were made.
        Console.WriteLine($"Replacements performed: {replacedCount}");

        // Save the modified document.
        doc.Save("OutputWithSmartTagsReplaced.docx");
    }
}
