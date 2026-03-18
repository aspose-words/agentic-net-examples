using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define file locations.
        string docsDir = @"C:\Docs\";
        string inputDocPath = Path.Combine(docsDir, "input.docx");   // source Word document
        string xmlMetadataPath = Path.Combine(docsDir, "metadata.xml"); // XML file to attach
        string outputPdfPath = Path.Combine(docsDir, "output.pdf"); // resulting PDF/A‑3u file

        // Load the Word document.
        Document doc = new Document(inputDocPath);

        // Insert the XML file as an embedded OLE object.
        // The progId "Package" is used for generic file attachments.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertOleObject(xmlMetadataPath, "Package", false, true, null);

        // Configure PDF save options:
        // - Set compliance to PDF/A‑3u.
        // - Embed attachments as document‑embedded files (no visible annotation).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.DocumentEmbeddedFiles
        };

        // Save the document as PDF/A‑3u with the XML attachment.
        doc.Save(outputPdfPath, pdfOptions);
    }
}
