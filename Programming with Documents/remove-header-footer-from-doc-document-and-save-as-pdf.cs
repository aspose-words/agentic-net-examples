using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndConvert
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\ResultDocument.pdf";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all headers and footers from the first (and only) section.
        // This clears the content but keeps the HeaderFooter objects, effectively unlinking them.
        doc.FirstSection.ClearHeadersFooters();

        // Save the modified document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
