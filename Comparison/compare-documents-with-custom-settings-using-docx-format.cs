using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class CompareDocumentsExample
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document revised = new Document("Revised.docx");

        // Set up comparison options with custom settings.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Use the revised document as the base for comparison.
            Target = ComparisonTargetType.New,
            // Do not ignore formatting changes.
            IgnoreFormatting = false,
            // Do not ignore case changes.
            IgnoreCaseChanges = false,
            // Do not ignore comments.
            IgnoreComments = false,
            // Do not ignore tables.
            IgnoreTables = false,
            // Do not ignore footnotes and endnotes.
            IgnoreFootnotes = false,
            // Do not ignore text boxes.
            IgnoreTextboxes = false,
            // Do not ignore headers and footers.
            IgnoreHeadersAndFooters = false,
            // Do not compare move operations.
            CompareMoves = false
        };

        // Advanced option: ignore differences in StructuredDocumentTag store item IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

        // Perform the comparison; revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result as a DOCX file using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        // Optional: set OOXML compliance level if required.
        // saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        original.Save("ComparisonResult.docx", saveOptions);
    }
}
