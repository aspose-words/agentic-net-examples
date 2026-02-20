using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        // The Document constructor automatically detects the file format.
        Document doc = new Document("Input.docx");

        // Create FindReplaceOptions to customize the replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        // Example: make the search case‑insensitive and match whole words only.
        options.MatchCase = false;
        options.FindWholeWordsOnly = true;

        // Perform the find‑and‑replace.
        // Replace every occurrence of "oldText" with "newText" using the specified options.
        doc.Range.Replace("oldText", "newText", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
