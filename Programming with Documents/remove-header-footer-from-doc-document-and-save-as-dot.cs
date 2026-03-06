using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeaderFooterAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Remove all header and footer contents from each section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Prepare save options for the DOT (Word template) format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);

        // Save the modified document as a DOT file.
        doc.Save("OutputTemplate.dot", saveOptions);
    }
}
