using System;
using Aspose.Words;
using Aspose.Words.Notes;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Load the source DOCM document.
        Document doc = new Document("SourceDocument.docm");

        // Index of the footnote to delete (0‑based). Adjust as needed.
        int footnoteIndexToDelete = 1;

        // Retrieve the footnote node. The third parameter 'true' searches recursively.
        Footnote footnote = (Footnote)doc.GetChild(NodeType.Footnote, footnoteIndexToDelete, true);

        // If the footnote exists, remove it from its parent.
        if (footnote != null)
        {
            footnote.Remove();
        }

        // Save the modified document as a DOTM template.
        doc.Save("ResultDocument.dotm");
    }
}
