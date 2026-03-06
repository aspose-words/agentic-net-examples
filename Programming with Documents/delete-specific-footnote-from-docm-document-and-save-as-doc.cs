using System;
using Aspose.Words;
using Aspose.Words.Notes;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Load the DOCM document. The file extension determines the format.
        Document doc = new Document("InputDocument.docm");

        // Retrieve all footnote nodes in the document (including those inside sections).
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Example: delete the footnote whose reference mark equals "2".
        // Adjust the condition as needed (e.g., by index, text content, etc.).
        foreach (Footnote fn in footnotes)
        {
            // You can also use fn.GetText() or fn.IndexOf(...) to identify the target.
            if (fn.ReferenceMark == "2")   // replace with your own criteria
            {
                fn.Remove();               // Removes the footnote from its parent.
                break;                     // Exit after deleting the desired footnote.
            }
        }

        // Save the modified document as a DOC file. The format is inferred from the extension.
        doc.Save("OutputDocument.doc");
    }
}
