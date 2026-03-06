using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Comparing;

class PdfComparisonExample
{
    static void Main()
    {
        // Paths to the original and edited PDF files.
        string originalPdfPath = @"C:\Docs\Original.pdf";
        string editedPdfPath = @"C:\Docs\Edited.pdf";

        // Load the PDF documents into Aspose.Words Document objects.
        Document docOriginal = new Document(originalPdfPath);
        Document docEdited = new Document(editedPdfPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");
        }

        // Set up comparison options (optional – customize as needed).
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: ignore formatting changes.
            IgnoreFormatting = true,
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF with revisions highlighted.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Ensure revisions are visible in the output PDF.
            UpdateFields = true
        };

        string resultPdfPath = @"C:\Docs\ComparisonResult.pdf";
        docOriginal.Save(resultPdfPath, saveOptions);
    }
}
