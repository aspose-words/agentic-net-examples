using System;
using Aspose.Words;

class RemoveHeaderFooter
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Remove headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Save the modified document as DOCX.
        doc.Save("OutputDocument.docx", SaveFormat.Docx);
    }
}
