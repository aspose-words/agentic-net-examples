using System;
using Aspose.Words;

namespace DocumentComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Load the original document from the file system.
            Document originalDoc = new Document("OriginalDocument.doc");

            // Load the document to compare against.
            Document editedDoc = new Document("EditedDocument.doc");

            // Ensure both documents have no revisions before performing the comparison.
            if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
            {
                // Compare the documents. All differences will be recorded as revisions in originalDoc.
                originalDoc.Compare(editedDoc, "JD", DateTime.Now);
            }

            // Save the result. The saved file will contain the revisions (tracked changes).
            originalDoc.Save("ComparisonResult.docx");
        }
    }
}
