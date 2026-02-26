using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFootersToPlainText
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path to the destination plain‑text file.
        string outputPath = @"C:\Docs\HeadersFooters.txt";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure text save options.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Export all headers and footers after the body text.
        // Options: None, PrimaryOnly, AllAtEnd.
        saveOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd;

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, saveOptions);

        // Optional: display the exported text in the console.
        string exportedText = File.ReadAllText(outputPath);
        Console.WriteLine("Exported plain‑text content:");
        Console.WriteLine(exportedText);
    }
}
