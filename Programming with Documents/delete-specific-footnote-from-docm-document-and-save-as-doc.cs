using System;
using Aspose.Words;
using Aspose.Words.Notes;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourcePath = @"C:\Docs\SourceDocument.docm";

        // Path for the resulting DOC file.
        string outputPath = @"C:\Docs\ResultDocument.doc";

        // Load the DOCM document. The constructor determines the format automatically.
        Document doc = new Document(sourcePath);

        // Define the reference mark of the footnote you want to delete.
        // If you want to delete by index, you can use the footnote's position in the collection instead.
        string targetReferenceMark = "RefMarkToDelete";

        // Retrieve all footnote nodes in the document.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Iterate through the collection to find the matching footnote.
        foreach (Footnote footnote in footnotes)
        {
            // Check if this footnote is a regular footnote (not an endnote) and matches the reference mark.
            if (footnote.FootnoteType == FootnoteType.Footnote && footnote.ReferenceMark == targetReferenceMark)
            {
                // Remove the footnote from its parent paragraph.
                footnote.Remove();
                // Exit after deleting the first matching footnote.
                break;
            }
        }

        // Save the modified document as a DOC file.
        // The Save method infers the format from the file extension.
        doc.Save(outputPath);
    }
}
