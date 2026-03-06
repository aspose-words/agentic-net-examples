using System;
using Aspose.Words;
using Aspose.Words.Saving;

class CompareDocumentsToPdf
{
    static void Main()
    {
        // Load the original document.
        Document docOriginal = new Document("Original.docx");

        // Load the edited document.
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the comparison result as a PDF file.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        docOriginal.Save("ComparisonResult.pdf", pdfOptions);
    }
}
