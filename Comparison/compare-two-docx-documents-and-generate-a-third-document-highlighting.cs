using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Set up comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level (you can use CharLevel for character‑level tracking).
            Granularity = Granularity.WordLevel,

            // Include all element types in the comparison.
            CompareMoves               = true,
            IgnoreFormatting           = false,
            IgnoreCaseChanges          = false,
            IgnoreComments             = false,
            IgnoreTables               = false,
            IgnoreFields               = false,
            IgnoreFootnotes            = false,
            IgnoreTextboxes            = false,
            IgnoreHeadersAndFooters    = false,

            // Use the original document as the base (show changes in the original).
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. The revisions (highlighted differences) are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document that contains the highlighted differences.
        docOriginal.Save("ComparisonResult.docx");
    }
}
