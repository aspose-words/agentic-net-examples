using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for PDF/A‑3u conversion.");

        // Create an XML metadata file that will be attached to the PDF.
        string xmlPath = Path.Combine(artifactsDir, "metadata.xml");
        File.WriteAllText(xmlPath, "<metadata><author>John Doe</author></metadata>");

        // Embed the XML file as an OLE object. The progId "Package" works for generic files.
        builder.InsertOleObject(xmlPath, "Package", false, true, null);

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as PDF/A‑3u.
        string pdfPath = Path.Combine(artifactsDir, "output.pdf");
        doc.Save(pdfPath, saveOptions);

        // Verify that the PDF file was created and is not empty.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the PDF/A‑3u file.");
    }
}
