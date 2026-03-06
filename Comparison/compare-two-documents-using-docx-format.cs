using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original document from disk.
        Document docOriginal = new Document("Original.docx");

        // Load the edited document that we want to compare against.
        Document docEdited = new Document("Edited.docx");

        // Both documents must not contain any revisions before we start the comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. All differences are recorded as revisions
            // in the original document (docOriginal).
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the result – the original document now contains revision marks
        // that highlight the changes between the two files.
        docOriginal.Save("Compared.docx");
    }
}
