using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class ComparisonExample
{
    static void Main()
    {
        // Paths to the source documents.
        string dataDir = @"C:\Docs\";
        string originalPath = dataDir + "Original.docx";
        string revisedPath = dataDir + "Revised.docx";

        // Load the documents.
        Document original = new Document(originalPath);
        Document revised = new Document(revisedPath);

        // Configure comparison options to ignore all changes that should not appear in the PDF.
        CompareOptions compareOptions = new CompareOptions
        {
            // Ignore formatting differences (bold, italic, font, etc.).
            IgnoreFormatting = true,
            // Ignore changes in headers and footers.
            IgnoreHeadersAndFooters = true,
            // Ignore table content changes.
            IgnoreTables = true,
            // Ignore footnote/endnote modifications.
            IgnoreFootnotes = true,
            // Ignore comment additions/removals.
            IgnoreComments = true,
            // Ignore text box content changes.
            IgnoreTextboxes = true,
            // Ignore field changes (e.g., DATE, REF).
            IgnoreFields = true,
            // Treat case changes as insignificant.
            IgnoreCaseChanges = true,
            // Do not track moved text as separate revisions.
            CompareMoves = false,
            // Track changes at character level (could also be WordLevel).
            Granularity = Granularity.CharLevel,
            // Use the revised document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Only non‑ignored differences will generate revisions.
        original.Compare(revised, "Reviewer", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF. The PDF will contain only the visible content,
        // because all ignored changes were excluded from the revision set.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        string outputPath = dataDir + "ComparisonResult.pdf";
        original.Save(outputPath, pdfOptions);
    }
}
