using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original document from a DOCX file.
            Document docOriginal = new Document("Original.docx");

            // Load the edited document from a DOCX file.
            Document docEdited = new Document("Edited.docx");

            // Ensure both documents have no revisions before comparison.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
            {
                // Compare the documents. All differences will be added as revisions to docOriginal.
                docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
            }

            // Save the comparison result (original document now contains revisions) to a new file.
            docOriginal.Save("ComparisonResult.docx");
        }
    }
}
