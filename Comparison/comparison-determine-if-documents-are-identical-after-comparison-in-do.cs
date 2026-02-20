using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DocumentComparer
{
    /// <summary>
    /// Determines whether two DOCX documents are identical after a comparison operation.
    /// Returns true if no revisions (differences) are detected.
    /// </summary>
    /// <param name="originalPath">Path to the first document (base document).</param>
    /// <param name="modifiedPath">Path to the second document (document to compare against).</param>
    /// <returns>True if documents are identical; otherwise false.</returns>
    public static bool AreDocumentsIdentical(string originalPath, string modifiedPath)
    {
        // Load the two documents.
        Document docOriginal = new Document(originalPath);
        Document docModified = new Document(modifiedPath);

        // Set up comparison options (default options are sufficient for a strict comparison).
        CompareOptions compareOptions = new CompareOptions
        {
            // Use the current document as the base for comparison.
            Target = ComparisonTargetType.Current,

            // Ensure all types of changes are considered.
            CompareMoves = true,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docModified, "Comparer", DateTime.Now, compareOptions);

        // If there are no revisions, the documents are identical.
        return docOriginal.Revisions.Count == 0;
    }

    // Example usage.
    public static void Main()
    {
        string pathA = @"C:\Docs\DocumentA.docx";
        string pathB = @"C:\Docs\DocumentB.docx";

        bool identical = AreDocumentsIdentical(pathA, pathB);
        Console.WriteLine(identical
            ? "The documents are identical."
            : "The documents differ.");
    }
}
