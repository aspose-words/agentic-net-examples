using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFooters
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Clear headers and footers in every section of the document.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters(); // Removes text but keeps the objects linked to previous sections.
        }

        // Save the modified document back to DOC format.
        doc.Save("OutputDocument.doc", SaveFormat.Doc);
    }
}
