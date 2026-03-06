using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove all headers and footers from each section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Save the document without headers/footers as a DOC file.
        doc.Save("Output.doc");
    }
}
