using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited DOCX documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Configure comparison options to ignore all types of changes.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,                 // Do not track moved text.
            IgnoreFormatting = true,              // Ignore formatting differences.
            IgnoreCaseChanges = true,             // Ignore case changes.
            IgnoreComments = true,                // Ignore comment changes.
            IgnoreTables = true,                  // Ignore table content changes.
            IgnoreFields = true,                  // Ignore field changes.
            IgnoreFootnotes = true,               // Ignore footnote/endnote changes.
            IgnoreTextboxes = true,               // Ignore textbox content changes.
            IgnoreHeadersAndFooters = true,       // Ignore header/footer changes.
            Target = ComparisonTargetType.New    // Use the edited document as the base.
        };

        // Perform the comparison. No revisions will be generated because all changes are ignored.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the resulting document (it will contain no revision marks).
        docOriginal.Save("Result.docx");
    }
}
