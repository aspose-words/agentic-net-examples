using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string pattern = "_FullName_";
        string replacement = "John Doe";

        // Configure find/replace options (case‑insensitive, whole‑word match).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement in each section's body.
        foreach (Section section in doc.Sections)
        {
            section.Body.Range.Replace(pattern, replacement, options);
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
