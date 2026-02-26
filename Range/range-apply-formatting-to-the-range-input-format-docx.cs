using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class ApplyFormattingToRange
{
    static void Main()
    {
        // 1. Load the source document (replace the path with the actual file location).
        Document doc = new Document(@"C:\Input\SourceDocument.docx");

        // 2. Define the text pattern to format. In this example we will bold every occurrence of the word "Important".
        string pattern = "Important";

        // 3. Configure find‑replace options to apply bold formatting.
        //    The FindReplaceOptions.ApplyFont property returns a Font object that can be modified.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ApplyFont.Bold = true;            // Apply bold style.
        options.MatchCase = false;               // Case‑insensitive search.
        options.FindWholeWordsOnly = true;       // Replace whole words only.

        // 4. Perform the replace operation. The replacement string is the same as the pattern,
        //    but the formatting defined in the options will be applied to each match.
        doc.Range.Replace(pattern, pattern, options);

        // 5. Save the modified document to a new file.
        doc.Save(@"C:\Output\FormattedDocument.docx", SaveFormat.Docx);
    }
}
