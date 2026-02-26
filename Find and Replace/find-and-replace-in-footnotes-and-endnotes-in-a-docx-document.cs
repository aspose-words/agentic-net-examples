using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Configure find/replace options.
        // Setting IgnoreFootnotes to false (default) ensures that footnotes and endnotes are included in the search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFootnotes = false
        };

        // Replace the target text throughout the entire document, including footnotes and endnotes.
        int replacementsMade = doc.Range.Replace("old text", "new text", options);

        // Optional: If you need to replace only inside footnotes/endnotes, iterate their ranges.
        // foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        // {
        //     footnote.Range.Replace("old text", "new text");
        // }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
