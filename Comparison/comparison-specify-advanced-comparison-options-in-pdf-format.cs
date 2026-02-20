using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document docOriginal = new Document("Original.docx");
        Document docRevised = new Document("Revised.docx");

        // Configure comparison options.
        // - Track changes at the word level.
        // - Do not ignore formatting or comments.
        // - Use the revised document as the base for comparison.
        // - Enable advanced options to ignore DrawingML unique IDs but consider SDT store item IDs.
        CompareOptions compareOptions = new CompareOptions
        {
            Granularity = Granularity.WordLevel,
            IgnoreFormatting = false,
            IgnoreComments = false,
            Target = ComparisonTargetType.New
        };
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;
        compareOptions.AdvancedOptions.IgnoreStoreItemId = false;

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docRevised, "Reviewer", DateTime.Now, compareOptions);

        // Set PDF save options.
        // - Save as PDF/A-2u (ISO 32000‑2) for long‑term preservation.
        // - Embed all fonts to ensure the PDF looks the same on any machine.
        // - Update fields before saving.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            EmbedFullFonts = true,
            UpdateFields = true
        };

        // Save the compared document as a PDF file.
        docOriginal.Save("ComparedResult.pdf", pdfOptions);
    }
}
