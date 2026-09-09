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
        builder.Writeln("This document will be saved as PDF/A‑3u with an attached XML metadata file.");

        // Prepare a sample XML metadata file.
        const string xmlFileName = "metadata.xml";
        File.WriteAllText(xmlFileName,
            "<metadata>" +
            "  <author>John Doe</author>" +
            "  <creationDate>2026-08-30</creationDate>" +
            "</metadata>");

        // Embed the XML file as an OLE object so it becomes an attachment in the PDF.
        // The last parameter (null) means no icon is displayed.
        builder.InsertOleObject(xmlFileName, "application/xml", false, true, null);

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException($"The file '{pdfFileName}' was not created.");

        Console.WriteLine($"PDF/A‑3u file '{pdfFileName}' created successfully with attached XML metadata.");
    }
}
