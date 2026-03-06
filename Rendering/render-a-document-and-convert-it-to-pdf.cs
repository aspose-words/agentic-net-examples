using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document from a file.
        Document doc = new Document("Input.docx");

        // Ensure the page layout is up‑to‑date before rendering.
        doc.UpdatePageLayout();

        // Save the document as PDF using the default save options.
        doc.Save("Output.pdf");
    }
}
