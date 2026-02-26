using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Create XPS save options (defaults to XPS format).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        xpsOptions.SaveFormat = SaveFormat.Xps; // Explicitly set for clarity.

        // Save the modified document as XPS.
        doc.Save("output.xps", xpsOptions);
    }
}
