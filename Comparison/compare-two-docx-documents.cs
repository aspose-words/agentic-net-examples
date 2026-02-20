using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document revised = new Document("Revised.docx");

        // Configure comparison options (customize as needed).
        CompareOptions options = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Do not ignore formatting differences.
            IgnoreFormatting = false,
            // Include comments, tables, footnotes, etc. in the comparison.
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            // Use the original document as the base for comparison.
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. The revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Save the document that now contains the revision marks.
        original.Save("ComparisonResult.docx");
    }
}
