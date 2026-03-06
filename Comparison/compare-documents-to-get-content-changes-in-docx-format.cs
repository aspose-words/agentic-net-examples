using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string dataDir = @"C:\Docs\";
        string originalPath = Path.Combine(dataDir, "Original.docx");
        string editedPath = Path.Combine(dataDir, "Edited.docx");

        // Load the documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited = new Document(editedPath);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options as needed.
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: track changes at the word level and do not ignore formatting.
                Granularity = Granularity.WordLevel,
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                CompareMoves = false,
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);
        }

        // Output details of each revision found.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document that now contains the tracked changes.
        string resultPath = Path.Combine(dataDir, "ComparedResult.docx");
        docOriginal.Save(resultPath);
    }
}
