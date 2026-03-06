using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Set up PostScript save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the document as a PostScript file.
        doc.Save("Output.ps", psOptions);
    }
}
