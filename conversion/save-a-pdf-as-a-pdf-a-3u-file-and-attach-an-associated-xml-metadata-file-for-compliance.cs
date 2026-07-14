using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple XML metadata file.
        const string xmlFileName = "metadata.xml";
        const string xmlContent = "<metadata><author>John Doe</author></metadata>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Build a basic Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for PDF/A-3u compliance.");

        // Embed the XML file as an OLE object; it will become an attachment in the PDF.
        builder.InsertOleObject(xmlFileName, "Package", false, true, null);

        // Configure PDF save options for PDF/A-3u and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as PDF/A-3u.
        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("The PDF/A-3u file was not created.");

        // Clean up temporary XML file (optional).
        if (File.Exists(xmlFileName))
            File.Delete(xmlFileName);
    }
}
