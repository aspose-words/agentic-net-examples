using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for PDF/A‑3u compliance.");

        // Create an XML metadata file that will be attached to the PDF.
        const string xmlFileName = "metadata.xml";
        File.WriteAllText(xmlFileName, "<metadata><author>John Doe</author></metadata>");

        // Insert the XML file as an OLE object so it can be embedded as an attachment.
        // "Package" is a generic OLE progID that works for arbitrary files.
        builder.InsertOleObject(xmlFileName, "Package", false, true, null);

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as a PDF/A‑3u file.
        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("The PDF/A‑3u file was not created.");

        // Clean up temporary files (optional).
        // File.Delete(xmlFileName);
    }
}
