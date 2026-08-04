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
        builder.Writeln("Sample document for PDF/A‑3u with XML attachment.");

        // Create a deterministic XML metadata file.
        const string xmlFileName = "metadata.xml";
        File.WriteAllText(xmlFileName, "<metadata><author>John Doe</author></metadata>");

        // Embed the XML file as an OLE object so it can be attached to the PDF.
        // The progId "Package" is a generic container for arbitrary files.
        builder.InsertOleObject(xmlFileName, "Package", false, true, null);

        // Configure PDF save options for PDF/A‑3u compliance and embed attachments.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        const string pdfFileName = "output_pdfa3u.pdf";
        doc.Save(pdfFileName, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("The PDF/A‑3u file was not created.");

        // Clean up temporary files (optional).
        File.Delete(xmlFileName);
    }
}
