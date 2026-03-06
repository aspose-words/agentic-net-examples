using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFootersToTxt
{
    static void Main()
    {
        // Load an existing DOCX document that contains headers and footers.
        // The Document constructor handles loading; no custom loading code is needed.
        Document doc = new Document("InputDocument.docx");

        // Create a TxtSaveOptions object to customize how the document is saved as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions();

        // Choose how headers and footers should appear in the exported TXT file.
        // Options: TxtExportHeadersFootersMode.None,
        //          TxtExportHeadersFootersMode.PrimaryOnly,
        //          TxtExportHeadersFootersMode.AllAtEnd
        txtOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;

        // Save the document as a .txt file using the configured options.
        // The Save method with a SaveOptions parameter follows the provided lifecycle rule.
        doc.Save("ExportedDocument.txt", txtOptions);

        // Optional: read the generated TXT file to verify its contents.
        string exportedText = File.ReadAllText("ExportedDocument.txt");
        Console.WriteLine("Exported TXT content:");
        Console.WriteLine(exportedText);
    }
}
