using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find-and-replace options: case‑sensitive and whole‑word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Replace all occurrences of the target text with the replacement text.
        doc.Range.Replace("oldText", "newText", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
