using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Configure find/replace options (case‑insensitive, whole‑word match).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement in each section's body.
        foreach (Section section in doc.Sections)
        {
            // The Range of the section's body contains all paragraphs, tables, etc.
            section.Body.Range.Replace(findText, replaceText, options);
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
