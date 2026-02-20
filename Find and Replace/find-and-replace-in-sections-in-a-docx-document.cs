using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceInSections
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        string findText = @"\[\[PLACEHOLDER\]\]"; // Regex pattern for [[PLACEHOLDER]]
        string replaceText = "New Text";

        // Set up find/replace options (optional, can customize behavior here).
        FindReplaceOptions options = new FindReplaceOptions();
        options.MatchCase = false;               // Case‑insensitive search.
        options.FindWholeWordsOnly = false;      // Allow matches inside larger words.

        // Iterate through each section and perform the replacement within that section's range.
        foreach (Section section in doc.Sections)
        {
            // Replace all occurrences of the pattern in the current section.
            section.Range.Replace(new Regex(findText), replaceText, options);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
