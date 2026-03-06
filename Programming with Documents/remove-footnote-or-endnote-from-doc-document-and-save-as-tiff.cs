using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.docx");

        // Remove all footnotes and endnotes.
        // Footnote nodes include both footnotes and endnotes.
        var footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        for (int i = footnoteNodes.Count - 1; i >= 0; i--)
        {
            Footnote footnote = (Footnote)footnoteNodes[i];
            footnote.Remove();
        }

        // Optionally remove footnote and endnote separators.
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        footnoteSeparator?.FirstParagraph?.FirstChild?.Remove();

        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        endnoteSeparator?.FirstParagraph?.FirstChild?.Remove();

        // Save the modified document as a TIFF image.
        doc.Save("output.tiff", SaveFormat.Tiff);
    }
}
