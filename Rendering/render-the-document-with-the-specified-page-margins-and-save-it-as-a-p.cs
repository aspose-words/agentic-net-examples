using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Apply a preset margin (e.g., Narrow) to every section in the document.
        foreach (Section section in doc.Sections)
        {
            section.PageSetup.Margins = Margins.Narrow;
        }

        // Save the document as a PDF file.
        doc.Save("Output.pdf");
    }
}
