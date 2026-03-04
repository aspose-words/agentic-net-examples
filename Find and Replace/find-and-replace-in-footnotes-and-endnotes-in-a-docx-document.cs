using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FootnoteEndnoteFindReplace
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create FindReplaceOptions and set IgnoreFootnotes to false
        // so that the replace operation also processes footnotes and endnotes.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFootnotes = false // false = include footnotes/endnotes in the search
        };

        // Define the text to find and its replacement.
        string findText = "Lorem ipsum";
        string replaceText = "Replaced Lorem ipsum";

        // Perform the find-and-replace across the whole document, including footnotes/endnotes.
        doc.Range.Replace(findText, replaceText, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
