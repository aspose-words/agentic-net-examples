using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Set the built‑in Title property – this value will be used for the PDF window title bar.
        doc.BuiltInDocumentProperties.Title = "My Document Title";

        // Add some visible content so the document is not empty.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Configure PDF save options to display the document title in the PDF title bar.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DisplayDocTitle = true
        };

        // Save the document as a PDF using the specified options.
        doc.Save("DocumentWithTitle.pdf", pdfOptions);
    }
}
