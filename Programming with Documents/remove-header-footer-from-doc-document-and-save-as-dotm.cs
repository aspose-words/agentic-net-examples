using System;
using Aspose.Words;

class RemoveHeadersFootersAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Clear the contents of all headers and footers in every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Save the modified document as a macro‑enabled template (DOTM).
        doc.Save("OutputTemplate.dotm", SaveFormat.Dotm);
    }
}
