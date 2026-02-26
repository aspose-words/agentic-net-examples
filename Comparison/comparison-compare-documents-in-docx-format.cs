using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the original document.
        Document docOriginal = new Document("Original.docx");

        // Load the document to compare against.
        Document docEdited = new Document("Edited.docx");

        // Documents must not contain revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the two documents.
            // The author name and the current date/time are recorded in the generated revisions.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the original document, now containing revision marks that represent the differences.
        docOriginal.Save("ComparisonResult.docx");
    }
}
