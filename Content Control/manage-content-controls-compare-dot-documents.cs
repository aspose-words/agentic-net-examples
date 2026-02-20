using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the two documents that will be compared.
        // These can be .dot template files or any other Word format.
        Document original = new Document("Original.dot");
        Document edited   = new Document("Edited.dot");

        // Set up comparison options.
        // - Track changes at the word level.
        // - Ignore formatting, comments, footnotes, and headers/footers.
        // - Use the edited document as the base (new) document.
        CompareOptions options = new CompareOptions
        {
            Granularity = Granularity.WordLevel,
            IgnoreFormatting = true,
            IgnoreComments = true,
            IgnoreFootnotes = true,
            IgnoreHeadersAndFooters = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions are added to the 'original' document.
        original.Compare(edited, "Comparer", DateTime.Now, options);

        // Iterate through all revisions and handle those that occur inside
        // Structured Document Tags (content controls).
        foreach (Revision rev in original.Revisions)
        {
            // Find the nearest ancestor StructuredDocumentTag, if any.
            StructuredDocumentTag sdt = rev.ParentNode?.GetAncestor(NodeType.StructuredDocumentTag) as StructuredDocumentTag;

            if (sdt != null)
            {
                // Example logic: accept revisions only in content controls with a specific tag.
                if (sdt.Tag == "Important")
                    rev.Accept();
                else
                    rev.Reject();
            }
        }

        // Save the resulting document with the applied revisions.
        original.Save("ComparedResult.docx");
    }
}
