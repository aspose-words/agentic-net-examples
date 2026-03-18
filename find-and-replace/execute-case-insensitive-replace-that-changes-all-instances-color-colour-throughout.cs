using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceColorWithColour
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path).
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Configure find/replace options for case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // ignore case when matching "color"
        };

        // Perform the replacement throughout the whole document.
        doc.Range.Replace("color", "colour", options);

        // Save the modified document (replace with your desired output path).
        doc.Save(@"C:\Docs\Output.docx");
    }
}
