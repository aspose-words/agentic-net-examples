using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the two DOCX files to compare.
        string filePathA = @"C:\Docs\DocumentA.docx";
        string filePathB = @"C:\Docs\DocumentB.docx";

        // Load the documents.
        Document docA = new Document(filePathA);
        Document docB = new Document(filePathB);

        // Ensure both documents have no revisions before comparison.
        if (docA.Revisions.Count != 0 || docB.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Set up comparison options.
        // Here we use the default options; adjust flags if you want to ignore specific differences.
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: ignore case changes, formatting, comments, etc.
            // IgnoreCaseChanges = true,
            // IgnoreFormatting = true,
            // IgnoreComments = true,
            // Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docA.
        docA.Compare(docB, "Comparer", DateTime.Now, compareOptions);

        // Determine if the documents are identical: no revisions means no differences.
        bool areIdentical = docA.Revisions.Count == 0;

        Console.WriteLine($"Documents are {(areIdentical ? "identical" : "different")}.");

        // Optionally, save the comparison result (with revisions) for inspection.
        if (!areIdentical)
        {
            string resultPath = @"C:\Docs\ComparisonResult.docx";
            docA.Save(resultPath);
            Console.WriteLine($"Comparison result saved to: {resultPath}");
        }
    }
}
