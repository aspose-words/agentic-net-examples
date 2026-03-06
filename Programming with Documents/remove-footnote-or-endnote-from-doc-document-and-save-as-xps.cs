using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveNotesAndSaveAsXps
{
    static void Main()
    {
        // Path to the folder that contains the input document.
        string dataDir = @"C:\Docs\"; // <-- change to your folder

        // Load the DOC/DOCX document.
        Document doc = new Document(dataDir + "Input.docx");

        // ----- Remove footnote separator (if present) -----
        // The separator is a paragraph that appears before the footnote collection.
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        if (footnoteSeparator?.FirstParagraph?.FirstChild != null)
        {
            // Remove the separator node.
            footnoteSeparator.FirstParagraph.FirstChild.Remove();
        }

        // ----- Remove endnote separator (if present) -----
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        if (endnoteSeparator?.FirstParagraph?.FirstChild != null)
        {
            // Remove the separator node.
            endnoteSeparator.FirstParagraph.FirstChild.Remove();
        }

        // Optionally, remove all footnote and endnote nodes themselves.
        // This loop removes each footnote/endnote from the document body.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            footnote.Remove();
        }

        // Save the modified document as XPS using XpsSaveOptions.
        XpsSaveOptions saveOptions = new XpsSaveOptions(); // default XPS format
        doc.Save(dataDir + "Output.xps", saveOptions);
    }
}
