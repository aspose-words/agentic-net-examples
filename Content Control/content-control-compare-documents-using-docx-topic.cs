using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the source documents and the output file.
        string docsFolder = @"C:\Docs\";
        string originalPath = Path.Combine(docsFolder, "Original.docx");
        string editedPath   = Path.Combine(docsFolder, "Edited.docx");
        string resultPath   = Path.Combine(docsFolder, "ComparisonResult.docx");

        // Load the two documents that will be compared.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure that both documents do not contain any revisions before comparison.
        // If revisions exist, the Compare method will throw an exception.
        if (docOriginal.Revisions.Count > 0 || docEdited.Revisions.Count > 0)
        {
            docOriginal.Revisions.AcceptAll();
            docEdited.Revisions.AcceptAll();
        }

        // Configure comparison options (optional). Here we compare all elements and track changes at the word level.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,                 // Do not generate move revisions.
            IgnoreFormatting = false,             // Track formatting changes.
            IgnoreCaseChanges = false,            // Track case changes.
            IgnoreComments = false,               // Track comment changes.
            IgnoreTables = false,                 // Track table changes.
            IgnoreFields = false,                 // Track field changes.
            IgnoreFootnotes = false,              // Track footnote/endnote changes.
            IgnoreTextboxes = false,              // Track textbox changes.
            IgnoreHeadersAndFooters = false,      // Track header/footer changes.
            Granularity = Granularity.WordLevel,  // Track changes by word.
            Target = ComparisonTargetType.Current // Use the original document as the base.
        };

        // Perform the comparison. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the document that now contains revision marks showing the differences.
        docOriginal.Save(resultPath);
    }
}
