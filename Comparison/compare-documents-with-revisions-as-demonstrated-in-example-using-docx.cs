using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and edited DOCX files.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before performing a comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options (all change types are tracked).
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

            // Compare the documents. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now, compareOptions);
        }

        // Save the result (original document now contains revision marks).
        docOriginal.Save("ComparedResult.docx");
    }
}
