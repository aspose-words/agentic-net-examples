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

        // Create a new blank Word document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with embedded XML metadata.");

        // Insert the XML file as an OLE object. This object will be saved as an attachment in the PDF/A‑3u file.
        // Parameters: file path, progId ("Package" works for generic files), isLink = false, isIcon = false, iconPath = null.
        builder.InsertOleObject(xmlFileName, "Package", false, false, null);

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as PDF/A‑3u.
        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, pdfOptions);

        // Verify that the PDF file was created and is not empty.
        if (!File.Exists(pdfFileName) || new FileInfo(pdfFileName).Length == 0)
        {
            throw new InvalidOperationException("The PDF/A‑3u file was not created successfully.");
        }
    }
}
