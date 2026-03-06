using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsTxt
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the plain‑text file will be saved.
        string outputPath = @"C:\Docs\Result.txt";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Optionally clear the header/footer contents of the first (and only) section.
        // This step is not strictly required if we set ExportHeadersFootersMode to None,
        // but it demonstrates how to remove the content programmatically.
        doc.FirstSection.ClearHeadersFooters();

        // Configure text save options to exclude headers and footers.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, txtOptions);
    }
}
