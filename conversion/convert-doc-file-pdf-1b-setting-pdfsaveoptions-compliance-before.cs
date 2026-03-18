using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToPdfA1b
{
    static void Main()
    {
        // Load the source DOC file.
        // Replace "input.doc" with the actual path to your DOC document.
        Document doc = new Document("input.doc");

        // Create PDF save options and set the compliance level to PDF/A‑1b.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF/A‑1b file.
        // Replace "output.pdf" with the desired output file path.
        doc.Save("output.pdf", saveOptions);
    }
}
