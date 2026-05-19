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
        builder.Writeln("Sample document for PDF/A‑3u conversion.");

        // Create a local XML file that will be embedded as metadata.
        const string xmlFileName = "metadata.xml";
        File.WriteAllText(xmlFileName, "<metadata><author>John Doe</author></metadata>");

        // Embed the XML file into the Word document as an OLE object.
        // The fourth parameter (isLink) is false to embed the file, and the fifth (isIconic) is false.
        builder.InsertOleObject(xmlFileName, "Package", false, false, null);

        // Save the Word document to a temporary DOCX file (required by the lifecycle rules).
        const string docxFileName = "input.docx";
        doc.Save(docxFileName, SaveFormat.Docx);

        // Load the saved DOCX document.
        Document loadedDoc = new Document(docxFileName);

        // Configure PDF/A‑3u save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/A‑3u.
            Compliance = PdfCompliance.PdfA3u,
            // Embed OLE objects as PDF attachments (annotations).
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as a PDF/A‑3u file with the embedded XML attachment.
        const string pdfFileName = "output.pdf";
        loadedDoc.Save(pdfFileName, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("The PDF/A‑3u file was not created.");

        // Optional cleanup (comment out if you want to inspect the files).
        // File.Delete(xmlFileName);
        // File.Delete(docxFileName);
        // File.Delete(pdfFileName);
    }
}
