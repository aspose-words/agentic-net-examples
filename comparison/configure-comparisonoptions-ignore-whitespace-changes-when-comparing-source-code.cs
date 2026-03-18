using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documentation files
        Document docOriginal = new Document("OriginalDocumentation.docx");
        Document docEdited = new Document("EditedDocumentation.docx");

        // Configure comparison options to ignore whitespace/formatting changes
        CompareOptions compareOptions = new CompareOptions
        {
            // Ignoring formatting treats whitespace differences as non‑significant
            IgnoreFormatting = true,
            // Keep other options at their defaults (explicitly set for clarity)
            CompareMoves = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the result document containing revisions
        docOriginal.Save("ComparisonResult.docx");
    }
}
