using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document. Replace the path with your source file.
        Document doc = new Document("Input.docx");

        // Create a TxtSaveOptions object to control plain‑text export.
        TxtSaveOptions txtOptions = new TxtSaveOptions();

        // Enable export of headers and footers.
        // Options: PrimaryOnly, AllAtEnd, or None.
        txtOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;

        // Save the document as a .txt file with the specified options.
        doc.Save("Output.txt", txtOptions);
    }
}
