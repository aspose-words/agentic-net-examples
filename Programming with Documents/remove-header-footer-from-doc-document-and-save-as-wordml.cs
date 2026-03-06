using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Remove headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            // Clears the text of all headers and footers while keeping the objects.
            section.ClearHeadersFooters();
        }

        // Prepare save options for WordML (Word 2003 XML) format.
        WordML2003SaveOptions saveOptions = new WordML2003SaveOptions
        {
            // Explicitly set the format to WordML (optional, default is WordML).
            SaveFormat = SaveFormat.WordML
        };

        // Save the modified document as WordML.
        doc.Save("Output.xml", saveOptions);
    }
}
