using System;
using Aspose.Words;
using Aspose.Words.Notes;

class RemoveFootnotesAndEndnotes
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "input.doc";

        // Path for the resulting DOT file.
        string outputPath = "output.dot";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Remove the footnote separator (if it exists).
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        if (footnoteSeparator?.FirstParagraph?.FirstChild != null)
            footnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Remove the endnote separator (if it exists).
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        if (endnoteSeparator?.FirstParagraph?.FirstChild != null)
            endnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Save the modified document as a DOT template.
        doc.Save(outputPath, SaveFormat.Dot);
    }
}
