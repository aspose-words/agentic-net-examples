using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Save the result which now contains revision marks.
        docOriginal.Save("Compared.docx");
    }
}
