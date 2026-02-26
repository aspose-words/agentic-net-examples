using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsWordML
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Clear headers and footers in every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Prepare WordML save options (Word 2003 XML format).
        WordML2003SaveOptions saveOptions = new WordML2003SaveOptions
        {
            // Explicitly set the format to WordML (optional, default is WordML).
            SaveFormat = SaveFormat.WordML
        };

        // Save the modified document as WordML.
        doc.Save("output.xml", saveOptions);
    }
}
