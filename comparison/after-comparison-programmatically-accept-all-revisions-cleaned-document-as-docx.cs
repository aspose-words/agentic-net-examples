using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documents.
        Document original = new Document("Original.docx");
        Document edited   = new Document("Edited.docx");

        // Compare the documents – revisions will be added to the original document.
        original.Compare(edited, "AuthorName", DateTime.Now);

        // Accept all revisions that were created by the comparison.
        original.AcceptAllRevisions();

        // Save the cleaned document (no revisions) as DOCX.
        original.Save("Cleaned.docx");
    }
}
