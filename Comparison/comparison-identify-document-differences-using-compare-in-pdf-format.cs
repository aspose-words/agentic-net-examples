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

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true,                         // Track moved content.
            Granularity = Granularity.WordLevel,         // Track changes by word.
            IgnoreFormatting = false,                    // Consider formatting changes.
            IgnoreComments = false,
            IgnoreFootnotes = false,
            IgnoreHeadersAndFooters = false,
            IgnoreTables = false,
            IgnoreTextboxes = false,
            IgnoreFields = false,
            IgnoreCaseChanges = false,
            Target = ComparisonTargetType.New           // Use the edited document as the base.
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the result as a PDF with revisions visible.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf17   // Use standard PDF 1.7 compliance.
        };

        docOriginal.Save("ComparisonResult.pdf", pdfOptions);
    }
}
