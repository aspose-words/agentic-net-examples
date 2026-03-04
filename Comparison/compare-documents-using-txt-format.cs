using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Comparing;

class CompareTxtDocuments
{
    static void Main()
    {
        // Path to the original and edited TXT files.
        const string originalPath = @"C:\Docs\Original.txt";
        const string editedPath   = @"C:\Docs\Edited.txt";

        // Load the TXT files into Document objects.
        // TxtLoadOptions allows us to specify how the plain‑text files are interpreted.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            // Example: detect hyperlinks while loading (optional).
            DetectHyperlinks = true
        };

        Document docOriginal = new Document(originalPath, loadOptions);
        Document docEdited   = new Document(editedPath,   loadOptions);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the two documents.
        // The original document will receive Revision objects describing the differences.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Output the number of revisions detected.
        Console.WriteLine($"Revisions found: {docOriginal.Revisions.Count}");

        // Optionally accept all revisions so that the original becomes identical to the edited document.
        docOriginal.Revisions.AcceptAll();

        // Save the merged result as a DOCX file.
        const string resultPath = @"C:\Docs\ComparisonResult.docx";
        docOriginal.Save(resultPath);
    }
}
