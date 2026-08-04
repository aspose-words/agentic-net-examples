using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample XML file that will be embedded as metadata.
        const string xmlFileName = "metadata.xml";
        File.WriteAllText(xmlFileName, "<metadata><author>John Doe</author></metadata>");

        // Create a blank Word document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with embedded XML metadata.");

        // Embed the XML file as an OLE object (attachment) in the document.
        // The progId "Package" is used for generic file attachments.
        builder.InsertOleObject(xmlFileName, "Package", false, true, null);

        // Configure PDF/A‑3u save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/A‑3u.
            Compliance = PdfCompliance.PdfA3u,
            // Embed attachments as annotations (required for PDF/A‑3).
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as a PDF/A‑3u file.
        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, pdfOptions);

        // Verify that the PDF file was created and is not empty.
        if (!File.Exists(pdfFileName) || new FileInfo(pdfFileName).Length == 0)
            throw new InvalidOperationException("The PDF/A‑3u file was not created successfully.");
    }
}
