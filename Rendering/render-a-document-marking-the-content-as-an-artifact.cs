using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Configure PDF save options to mark paragraph graphics (e.g., underlines, emphasis) as artifacts.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,               // Enable logical structure export.
            ExportParagraphGraphicsToArtifact = true,     // Mark paragraph graphics as artifacts.
            TextCompression = PdfTextCompression.None    // Optional: keep text uncompressed for easier inspection.
        };

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", saveOptions);
    }
}
