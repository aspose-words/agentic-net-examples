using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Simple string replacement (case‑insensitive by default).
        // Replaces all occurrences of the placeholder _FullName_ with "John Doe".
        int replacedCount = doc.Range.Replace("_FullName_", "John Doe");

        // Example of a regex replacement with additional options.
        // Replaces the whole word "OldCompany" with "NewCompany", ignoring case.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };
        doc.Range.Replace(new Regex(@"\bOldCompany\b"), "NewCompany", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
