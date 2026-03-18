using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareSpecificSections
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document(@"C:\Docs\Original.docx");
        Document docEdited   = new Document(@"C:\Docs\Edited.docx");

        // -----------------------------------------------------------------
        // Prepare comparison options so that only the main body text of the
        // first section is taken into account. All other element types are
        // ignored (tables, footnotes, headers/footers, etc.).
        // -----------------------------------------------------------------
        CompareOptions compareOptions = new CompareOptions
        {
            // Do not generate move revisions.
            CompareMoves = false,

            // Ignore formatting changes – we only care about the text itself.
            IgnoreFormatting = true,

            // Case‑insensitive comparison (optional, set to false to keep case).
            IgnoreCaseChanges = false,

            // Suppress revisions for all element types that are not plain text.
            IgnoreComments          = true,
            IgnoreTables            = true,
            IgnoreFields            = true,
            IgnoreFootnotes         = true,
            IgnoreTextboxes         = true,
            IgnoreHeadersAndFooters = true,

            // Use the edited document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // -----------------------------------------------------------------
        // To limit the comparison to a specific section we extract that
        // section from each document into a temporary document and compare
        // those temporary documents.
        // -----------------------------------------------------------------
        Document tempOriginal = ExtractSection(docOriginal, 0); // first section (index 0)
        Document tempEdited   = ExtractSection(docEdited,   0);

        // Perform the comparison using the configured options.
        tempOriginal.Compare(tempEdited, "Analyzer", DateTime.Now, compareOptions);

        // Save the comparison result – it will contain revisions only for the
        // first section's body text.
        tempOriginal.Save(@"C:\Docs\ComparisonResult.docx");
    }

    // Helper method that creates a new document containing only the section
    // at the specified index from the source document.
    private static Document ExtractSection(Document source, int sectionIndex)
    {
        // Clone the requested section to avoid modifying the original document.
        Section sectionClone = (Section)source.Sections[sectionIndex].Clone(true);

        // Create a new blank document.
        Document result = new Document();
        // Remove the default empty section that the constructor creates.
        result.RemoveAllChildren();

        // The cloned section is already a detached node, so we can simply
        // append it to the new document. No NodeImporter is required.
        result.AppendChild(sectionClone);

        return result;
    }
}
