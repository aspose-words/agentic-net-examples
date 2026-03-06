using System;
using Aspose.Words;

class RemoveHeadersFootersAndSaveAsDocm
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Iterate through all sections and clear their headers and footers.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Save the modified document as a macro‑enabled DOCM file.
        doc.Save("OutputDocument.docm", SaveFormat.Docm);
    }
}
