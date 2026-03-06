using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        Document doc = new Document("Input.doc");

        // Remove all header and footer contents from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Save the modified document as a macro‑enabled template (DOTM).
        doc.Save("Output.dotm", SaveFormat.Dotm);
    }
}
