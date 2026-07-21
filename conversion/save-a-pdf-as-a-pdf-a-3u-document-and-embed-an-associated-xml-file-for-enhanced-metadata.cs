using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary XML metadata file and the resulting PDF/A‑3u file.
        const string xmlPath = "metadata.xml";
        const string pdfPath = "outputPdfA3u.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple XML file that will be embedded in the PDF.
        // -----------------------------------------------------------------
        File.WriteAllText(xmlPath,
            "<metadata>" +
            "<author>John Doe</author>" +
            "<description>Sample PDF/A-3u with embedded XML attachment</description>" +
            "</metadata>");

        // -----------------------------------------------------------------
        // 2. Build a basic Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is saved as PDF/A-3u and contains an embedded XML file.");

        // -----------------------------------------------------------------
        // 3. Embed the XML file as an OLE object.
        //    The \"Package\" progID creates a generic file attachment.
        // -----------------------------------------------------------------
        builder.InsertOleObject(xmlPath, "Package", false, true, null);

        // -----------------------------------------------------------------
        // 4. Configure PDF save options for PDF/A‑3u compliance.
        //    Use the Annotations mode to embed the attachment (supported in this SDK version).
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations,
            ExportDocumentStructure = true
        };

        // -----------------------------------------------------------------
        // 5. Save the document as a PDF/A‑3u file.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, saveOptions);

        // -----------------------------------------------------------------
        // 6. Verify that the PDF/A‑3u file was created and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("The PDF/A-3u file was not created successfully.");

        // Optional cleanup of the temporary XML file.
        // File.Delete(xmlPath);
    }
}
