using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no existing revisions (required before comparison).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "AuthorInitials", DateTime.Now);

        // Optionally, accept all revisions to transform docOriginal into the edited version.
        // docOriginal.Revisions.AcceptAll();

        // Save the comparison result as a PDF file.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
