using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA3uWithXmlEmbedding
{
    static void Main()
    {
        // Create a temporary folder for the example files.
        string dataDir = Path.Combine(Path.GetTempPath(), "PdfA3uExample");
        Directory.CreateDirectory(dataDir);

        // Create a simple XML file to embed.
        string xmlPath = Path.Combine(dataDir, "metadata.xml");
        File.WriteAllText(xmlPath, "<metadata><author>John Doe</author></metadata>");

        // Create a new, empty document.
        Document doc = new Document();

        // Build a simple document body.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document for PDF/A‑3u with embedded XML metadata.");

        // Embed the XML file as an OLE object.
        // The progId "Package" is used for generic file embedding.
        builder.InsertOleObject(
            xmlPath,          // full path to the XML file
            "Package",        // OLE progId for generic file embedding
            false,            // do not link, embed the file data
            false,            // do not display as an icon
            null);            // no additional image for the icon

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.DocumentEmbeddedFiles
        };

        // Save the document as a PDF/A‑3u file with the XML file embedded.
        string pdfPath = Path.Combine(dataDir, "OutputPdfA3u.pdf");
        doc.Save(pdfPath, saveOptions);

        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
