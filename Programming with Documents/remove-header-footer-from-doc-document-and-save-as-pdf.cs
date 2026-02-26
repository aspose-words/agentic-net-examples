using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Remove all headers and footers from every section in the document.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Save the resulting document as a PDF file.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
