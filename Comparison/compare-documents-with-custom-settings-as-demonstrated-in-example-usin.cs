using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Set up custom comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves               = false,                     // Do not track moved text.
            IgnoreFormatting           = false,                     // Consider formatting changes.
            IgnoreCaseChanges          = false,                     // Case changes are significant.
            IgnoreComments             = false,                     // Compare comments.
            IgnoreTables               = false,                     // Compare table content.
            IgnoreFields               = false,                     // Compare fields.
            IgnoreFootnotes            = false,                     // Compare footnotes/endnotes.
            IgnoreTextboxes            = false,                     // Compare text inside text boxes.
            IgnoreHeadersAndFooters    = false,                     // Compare headers/footers.
            Target                     = ComparisonTargetType.New, // Use the edited document as the base.
            Granularity                = Granularity.WordLevel    // Track changes at word level.
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document with a specific OOXML compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        docOriginal.Save("ComparedResult.docx", saveOptions);
    }
}
