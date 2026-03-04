using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Path to the folder that contains the source documents.
        string docsPath = @"C:\Docs\"; // <-- adjust as needed

        // Load the original and the edited documents.
        Document original = new Document(docsPath + "Original.docx");
        Document edited   = new Document(docsPath + "Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Create comparison options (default settings can be used or customized here).
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: track changes at the word level.
                Granularity = Granularity.WordLevel,
                // Example: ignore case changes.
                IgnoreCaseChanges = true,
                // Example: set the target document for comparison.
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions will be added to the original document.
            original.Compare(edited, "Comparer", DateTime.Now, compareOptions);
        }

        // Save the comparison result in WORDML (Word 2003 XML) format.
        original.Save(docsPath + "ComparisonResult.xml", SaveFormat.WordML);
    }
}
