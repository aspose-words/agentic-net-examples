using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to the source DOCM files.
            const string originalPath = @"C:\Docs\Original.docm";
            const string editedPath   = @"C:\Docs\Edited.docm";

            // Load the original and edited documents.
            Document docOriginal = new Document(originalPath);
            Document docEdited   = new Document(editedPath);

            // Ensure both documents have no tracked revisions before comparison.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

            // Perform the comparison. The original document will receive Revision objects
            // describing the differences found in the edited document.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

            // Optional: accept all revisions to transform the original into the edited version.
            // docOriginal.Revisions.AcceptAll();

            // Save the comparison result as a DOCM file.
            const string resultPath = @"C:\Docs\ComparisonResult.docm";
            docOriginal.Save(resultPath);
        }
    }
}
