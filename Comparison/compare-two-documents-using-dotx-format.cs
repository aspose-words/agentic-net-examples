using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDotxDocuments
{
    static void Main()
    {
        // Paths to the source DOTX files.
        string originalPath = @"C:\Docs\Original.dotx";
        string editedPath   = @"C:\Docs\Edited.dotx";

        // Load the two documents. The Document(string) constructor automatically detects the DOTX format.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure that neither document contains revisions before performing the comparison.
        // If revisions exist, the Compare method will throw an exception.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. The original document will receive Revision objects describing the differences.
        // Author name and timestamp are required parameters.
        docOriginal.Compare(docEdited, "Author", DateTime.Now);

        // Optionally, accept all revisions to transform the original into the edited version.
        // docOriginal.Revisions.AcceptAll();

        // Save the result back to DOTX format.
        string resultPath = @"C:\Docs\Compared.dotx";
        docOriginal.Save(resultPath);
    }
}
