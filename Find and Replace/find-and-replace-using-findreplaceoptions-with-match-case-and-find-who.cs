using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
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

        // Replace all occurrences of "oldWord" with "newWord" using the options.
        doc.Range.Replace("oldWord", "newWord", options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
