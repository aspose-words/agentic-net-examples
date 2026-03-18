using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceWholeWordsCaseInsensitive
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options:
        // - MatchCase = false  -> case‑insensitive search.
        // - FindWholeWordsOnly = true -> replace only whole word matches.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement across the whole document.
        // Replace every occurrence of "oldWord" with "newWord" using the options above.
        doc.Range.Replace("oldWord", "newWord", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
