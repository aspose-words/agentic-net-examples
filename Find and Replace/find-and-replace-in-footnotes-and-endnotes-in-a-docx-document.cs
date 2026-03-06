using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Configure find/replace options to include footnotes and endnotes.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // The default is false, but we set it explicitly for clarity.
            IgnoreFootnotes = false
        };

        // Perform the replacement across the whole document, including footnotes/endnotes.
        int replacements = doc.Range.Replace("OldText", "NewText", options);
        Console.WriteLine($"Replacements made: {replacements}");

        // Save the modified document.
        doc.Save(@"C:\Docs\Output.docx");
    }
}
