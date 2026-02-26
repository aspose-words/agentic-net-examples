using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the two DOCX files to compare.
        string pathOriginal = "Original.docx";
        string pathEdited   = "Edited.docx";

        // Load the documents from disk.
        Document docOriginal = new Document(pathOriginal);
        Document docEdited   = new Document(pathEdited);

        // Ensure both documents are revision‑free before comparison (required by Aspose.Words).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must not contain revisions before comparison.");

        // Perform the comparison. The original document will receive revisions for any differences.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // If no revisions were added, the documents are identical.
        bool areIdentical = docOriginal.Revisions.Count == 0;

        Console.WriteLine(areIdentical
            ? "The documents are identical."
            : "The documents differ.");

        // Optional: save the original document with revision markup for visual inspection.
        docOriginal.Save("ComparisonResult.docx");
    }
}
