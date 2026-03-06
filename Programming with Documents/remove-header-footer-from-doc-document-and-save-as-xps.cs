using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOC document.
        Document doc = new Document("input.doc");

        // Remove headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Create XPS save options (defaults to Xps format).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the modified document as XPS.
        doc.Save("output.xps", xpsOptions);
    }
}
