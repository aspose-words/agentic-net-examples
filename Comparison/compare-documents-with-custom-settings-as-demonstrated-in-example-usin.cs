using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the source documents and the output folder.
        string dataDir = @"C:\Data\";
        string artifactsDir = @"C:\Artifacts\";

        // Load the original document.
        Document docOriginal = new Document(dataDir + "Original.docx");

        // Load the edited document that we want to compare against.
        Document docEdited = new Document(dataDir + "Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Configure custom comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: ignore formatting changes and comments.
            IgnoreFormatting = true,
            IgnoreComments = true,
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Use the edited document as the base for comparison.
            Target = ComparisonTargetType.New,
            // Do not compare moved text.
            CompareMoves = false
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the resulting document that contains the tracked changes.
        docOriginal.Save(artifactsDir + "ComparisonResult.docx");
    }
}
