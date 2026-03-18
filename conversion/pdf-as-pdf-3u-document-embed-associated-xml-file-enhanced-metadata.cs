using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA3uWithXmlEmbedding
{
    static void Main()
    {
        // Path to the folder that contains the source XML file and where the PDF will be saved.
        string dataDir = @"C:\Data\";

        // Create a new, empty document.
        Document doc = new Document();

        // Build a simple document body.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document for PDF/A‑3u with embedded XML metadata.");

        // Embed the XML file as an OLE object. The last parameter (true) tells Aspose.Words
        // to embed the file data, which later can be saved as an attachment in the PDF.
        builder.InsertOleObject(
            Path.Combine(dataDir, "metadata.xml"), // full path to the XML file
            "text/xml",                           // MIME type of the embedded file
            false,                                // do not display as an icon
            true,                                 // embed the file data
            null);                                // no additional image for the icon

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.DocumentEmbeddedFiles
        };

        // Save the document as a PDF/A‑3u file with the XML file embedded.
        doc.Save(Path.Combine(dataDir, "OutputPdfA3u.pdf"), saveOptions);
    }
}
