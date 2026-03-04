using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class TxtDocumentComparison
{
    static void Main()
    {
        // Paths to the original and edited TXT files.
        string originalPath = Path.Combine(Environment.CurrentDirectory, "original.txt");
        string editedPath   = Path.Combine(Environment.CurrentDirectory, "edited.txt");

        // Load the TXT files into Aspose.Words Document objects using TxtLoadOptions.
        // This ensures the plain‑text files are interpreted correctly.
        Document docOriginal = new Document(originalPath, new TxtLoadOptions());
        Document docEdited   = new Document(editedPath,   new TxtLoadOptions());

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. All differences will be recorded as revisions in docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Output the number of revisions (differences) found.
        Console.WriteLine($"Total revisions detected: {docOriginal.Revisions.Count}");

        // List each revision with its type and the changed text.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}");
            Console.WriteLine($"Changed text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Optional: accept all revisions to make docOriginal identical to docEdited.
        // docOriginal.Revisions.AcceptAll();

        // Optional: save the comparison result as a DOCX file for visual inspection.
        // string resultPath = Path.Combine(Environment.CurrentDirectory, "ComparisonResult.docx");
        // docOriginal.Save(resultPath);
    }
}
