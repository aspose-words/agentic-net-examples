using System;
using Aspose.Words;
using Aspose.Words.Notes;

class RemoveNotesAndSaveAsDocm
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Remove the footnote separator (the line that separates footnotes from the main text).
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        // Ensure the separator has a child node before attempting to remove it.
        if (footnoteSeparator.FirstParagraph?.FirstChild != null)
            footnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Remove the endnote separator (the line that separates endnotes from the main text).
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        if (endnoteSeparator.FirstParagraph?.FirstChild != null)
            endnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Save the modified document as a macro‑enabled DOCM file.
        doc.Save("OutputDocument.docm");
    }
}
