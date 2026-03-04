using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocs
{
    static void Main()
    {
        // Load the original and edited DOCX documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Configure comparison options (adjust flags as needed).
        CompareOptions options = new CompareOptions
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
            Target = ComparisonTargetType.New // Show changes in the edited document.
        };

        // Perform the comparison; revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "JD", DateTime.Now, options);

        // Iterate through the generated revisions and output details.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText()}\"");
        }

        // Save the document that now contains the revision markup.
        docOriginal.Save("ComparedResult.docx");
    }
}
