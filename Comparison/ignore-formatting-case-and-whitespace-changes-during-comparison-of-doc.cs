using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options to ignore formatting, case changes and treat whitespace as insignificant.
        CompareOptions compareOptions = new CompareOptions
        {
            // Ignore any formatting differences (font, style, etc.).
            IgnoreFormatting = true,
            // Make the comparison case‑insensitive.
            IgnoreCaseChanges = true,
            // Track changes at the character level to catch whitespace modifications.
            Granularity = Granularity.CharLevel,
            // Use the edited document as the base for comparison (optional, default is Current).
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the result which contains the revisions (differences) after ignoring the specified aspects.
        docOriginal.Save("ComparisonResult.docx");
    }
}
