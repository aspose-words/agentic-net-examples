using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFooters
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create TxtSaveOptions and configure it to export headers and footers.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        // Export all headers and footers after the main body of each section.
        txtOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd;

        // Save the document as plain‑text, applying the configured options.
        string txtFilePath = "HeadersFooters.txt";
        doc.Save(txtFilePath, txtOptions);

        // Output the resulting text to the console (optional).
        Console.WriteLine(File.ReadAllText(txtFilePath));
    }
}
