using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original document (must have no revisions before comparison).
            Document docOriginal = new Document("Original.docx");

            // Load the document to compare against.
            Document docEdited = new Document("Edited.docx");

            // Ensure both documents are revision‑free; otherwise Compare will throw.
            if (docOriginal.Revisions.Count > 0 || docEdited.Revisions.Count > 0)
                throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

            // Perform the comparison. The original document will receive Revision objects
            // describing the differences found in the edited document.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

            // Save the result. The file will contain tracked changes (revisions).
            docOriginal.Save("ComparedResult.docx");
        }
    }
}
