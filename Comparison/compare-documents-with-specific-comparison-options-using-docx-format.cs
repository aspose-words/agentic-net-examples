using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original and the edited documents from the file system.
            Document docOriginal = new Document("Original.docx");
            Document docEdited   = new Document("Edited.docx");

            // Configure comparison options.
            // In this example we ignore formatting changes but track all other differences.
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreFormatting = true,                 // Do not treat formatting changes as revisions.
                IgnoreComments    = false,                // Include comment changes.
                IgnoreTables      = false,                // Include table changes.
                IgnoreFields      = false,                // Include field changes.
                IgnoreFootnotes   = false,                // Include footnote/endnote changes.
                IgnoreTextboxes   = false,                // Include textbox changes.
                IgnoreHeadersAndFooters = false,          // Include header/footer changes.
                Target = ComparisonTargetType.New      // Use the edited document as the base for comparison.
            };

            // Perform the comparison. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Save the resulting document (which now contains revision marks) in DOCX format.
            docOriginal.Save("ComparedResult.docx");
        }
    }
}
