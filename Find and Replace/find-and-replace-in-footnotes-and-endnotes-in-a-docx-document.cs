using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options to process footnotes and endnotes.
        FindReplaceOptions options = new FindReplaceOptions();
        options.IgnoreFootnotes = false; // false (default) means footnotes/endnotes are included.

        // Perform the find and replace across the main story, footnotes, and endnotes.
        // Example: replace every occurrence of "oldText" with "newText".
        doc.Range.Replace(new Regex("oldText"), "newText", options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
