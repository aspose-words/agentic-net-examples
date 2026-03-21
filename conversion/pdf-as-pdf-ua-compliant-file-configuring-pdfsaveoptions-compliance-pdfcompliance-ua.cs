using Aspose.Words;
using Aspose.Words.Saving;

class PdfUaExport
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, PDF/UA!");

        // Configure PDF save options for PDF/UA‑1 compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1
        };

        // Save the document as a PDF/UA‑compliant file.
        doc.Save("Output.pdf", pdfOptions);
    }
}
