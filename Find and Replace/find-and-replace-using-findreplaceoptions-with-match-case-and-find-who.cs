using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceExample
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options: case-sensitive and whole-word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        // Replace all occurrences of "Apple" with "Orange" respecting the options.
        doc.Range.Replace("Apple", "Orange", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
