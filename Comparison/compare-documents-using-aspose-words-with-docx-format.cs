using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the source documents and the output file.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";
        string resultPath   = "ComparedResult.docx";

        // Load the two documents that will be compared.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Aspose.Words requires that both documents have no revisions before a comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison.
        // The first document (docOriginal) will receive Revision objects that represent the differences.
        // "Comparer" is the author name that will appear on the revisions.
        // DateTime.Now supplies the timestamp for the revisions.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Output the number of revisions created for diagnostic purposes.
        Console.WriteLine($"Revisions created: {docOriginal.Revisions.Count}");

        // Save the original document now containing the revisions.
        docOriginal.Save(resultPath);
    }
}
