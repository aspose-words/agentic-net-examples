using System;
using Aspose.Words;
using Aspose.Words.Notes;

class DeleteSpecificFootnote
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Specify the zero‑based index of the footnote to delete.
        // For example, to delete the second footnote use index = 1.
        int footnoteIndex = 1;

        // Retrieve the footnote node. The third parameter (true) searches recursively.
        Footnote footnote = (Footnote)doc.GetChild(NodeType.Footnote, footnoteIndex, true);

        // If the footnote exists, remove it from its parent.
        if (footnote != null)
        {
            footnote.Remove();
        }

        // Save the modified document as a DOTM template.
        doc.Save("OutputDocument.dotm");
    }
}
