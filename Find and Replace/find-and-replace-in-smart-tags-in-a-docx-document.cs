using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class SmartTagFindReplace
{
    static void Main()
    {
        // Load the DOCX document that contains smart tags.
        Document doc = new Document(@"C:\Docs\InputWithSmartTags.docx");

        // Configure find/replace options to treat the content of StructuredDocumentTag (SmartTag)
        // as plain text, allowing the replace operation to work inside smart tags.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreStructuredDocumentTags = true   // Treat smart tag content as simple text.
        };

        // Perform the find-and-replace operation.
        // Example: replace the placeholder "_FullName_" with "John Doe".
        int replacements = doc.Range.Replace("_FullName_", "John Doe", options);

        Console.WriteLine($"Number of replacements made: {replacements}");

        // Save the modified document.
        doc.Save(@"C:\Docs\OutputWithSmartTagsReplaced.docx");
    }
}
