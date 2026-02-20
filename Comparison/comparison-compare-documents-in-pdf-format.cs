using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class PdfComparisonExample
{
    static void Main()
    {
        // Paths to the source PDF documents.
        string originalPdfPath = @"C:\Docs\Original.pdf";
        string editedPdfPath   = @"C:\Docs\Edited.pdf";

        // Load the PDF documents into Aspose.Words Document objects.
        Document originalDoc = new Document(originalPdfPath);
        Document editedDoc   = new Document(editedPdfPath);

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New,
            // Example: ignore formatting differences.
            IgnoreFormatting = true
        };

        // Perform the comparison. Revisions will be added to originalDoc.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            // Ensure the output complies with PDF 1.7.
            Compliance = PdfCompliance.Pdf17
        };

        string resultPdfPath = @"C:\Docs\ComparisonResult.pdf";
        originalDoc.Save(resultPdfPath, pdfSaveOptions);
    }
}
