using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the first PDF document.
        var loadOptions = new PdfLoadOptions(); // Use default load options.
        Document docOriginal = new Document(@"C:\Docs\Original.pdf", loadOptions);

        // Load the second PDF document to compare with.
        Document docEdited = new Document(@"C:\Docs\Edited.pdf", loadOptions);

        // Set up comparison options (customize as needed).
        var compareOptions = new CompareOptions
        {
            // Example: ignore formatting changes.
            IgnoreFormatting = true,
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Use the edited document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF.
        var saveOptions = new PdfSaveOptions
        {
            // Ensure the output is a standard PDF (not PDF/A).
            Compliance = PdfCompliance.Pdf17
        };
        docOriginal.Save(@"C:\Docs\ComparisonResult.pdf", saveOptions);
    }
}
