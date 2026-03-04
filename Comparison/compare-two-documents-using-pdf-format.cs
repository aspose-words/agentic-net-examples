using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparison
{
    static void Main()
    {
        // Load the original document.
        Document original = new Document("original.docx");

        // Load the edited document to compare against.
        Document edited = new Document("edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The revisions will be added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Save the comparison result as a PDF file.
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // Default options.
        original.Save("ComparisonResult.pdf", pdfOptions);
    }
}
