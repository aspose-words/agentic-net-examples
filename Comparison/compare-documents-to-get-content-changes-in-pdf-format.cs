using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparisonToPdf
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Compare the documents. The original document will receive revision marks.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the result as a PDF. Revisions are preserved in the PDF output.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
