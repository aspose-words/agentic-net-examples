using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFooters
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Iterate through all sections and clear their headers and footers.
        foreach (Section section in doc.Sections)
        {
            // This removes the text from all header/footer types but keeps the objects,
            // effectively unlinking them from any previous sections.
            section.ClearHeadersFooters();
        }

        // Save the modified document as a macro‑enabled DOCM file.
        // Using OoxmlSaveOptions to explicitly set the target format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        doc.Save("output.docm", saveOptions);
    }
}
