using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class CaseSensitiveReplaceExample
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Configure find-and-replace options to be case‑sensitive.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true // Enable case‑sensitive comparison.
        };

        // Perform the replace operation.
        // Replace all occurrences of "Apple" with "Orange" respecting case.
        doc.Range.Replace("Apple", "Orange", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
