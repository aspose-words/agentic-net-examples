using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // The Compare method requires that both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. All differences are added to docOriginal as revisions.
            // "Comparer" will appear as the author of the revisions, and the current time is used.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the resulting document (with revisions) as a PDF.
        // PdfSaveOptions can be customized if needed; here we use the defaults.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        docOriginal.Save("ComparisonResult.pdf", pdfOptions);
    }
}
