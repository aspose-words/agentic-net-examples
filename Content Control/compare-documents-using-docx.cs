using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents from disk.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents are free of revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Configure comparison options (optional – adjust flags as needed).
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Output the number of revisions detected.
        Console.WriteLine($"Revisions found: {docOriginal.Revisions.Count}");

        // Accept all revisions so docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document.
        docOriginal.Save("ComparedResult.docx");
    }
}
