using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the two TXT files to compare.
        string originalPath = "Original.txt";
        string editedPath   = "Edited.txt";

        // Load the TXT files into Document objects using TxtLoadOptions.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        Document docOriginal = new Document(originalPath, loadOptions);
        Document docEdited   = new Document(editedPath,   loadOptions);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions describing differences are added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Output the number of revisions (differences) found.
        Console.WriteLine($"Revisions found: {docOriginal.Revisions.Count}");

        // Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting merged document back to TXT format.
        docOriginal.Save("MergedResult.txt", SaveFormat.Text);
    }
}
