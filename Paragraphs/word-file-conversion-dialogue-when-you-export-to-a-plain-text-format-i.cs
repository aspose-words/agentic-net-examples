using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocxToPlainText
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SampleDocument.docx";

        // Path where the plain‑text file will be saved.
        string targetPath = @"C:\Docs\SampleDocument.txt";

        // Load the DOCX document from disk.
        Document doc = new Document(sourcePath);

        // Configure plain‑text save options.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Explicitly set the format to Text (optional – default is Text).
            SaveFormat = SaveFormat.Text,

            // Example: export only primary headers/footers.
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly,

            // Example: preserve page breaks in the output.
            ForcePageBreaks = true
        };

        // Save the document as plain text using the configured options.
        doc.Save(targetPath, txtOptions);
    }
}
