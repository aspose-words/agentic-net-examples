using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This document will be saved as PDF/A‑3u with an attached XML file.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the document we just created.
        Document doc = new Document("input.docx");

        // Create a sample XML metadata file.
        string xmlContent = "<metadata><author>John Doe</author><created>2026-07-11</created></metadata>";
        File.WriteAllText("metadata.xml", xmlContent);

        // Embed the XML file as an OLE object (it will become an attachment in PDF/A‑3u).
        DocumentBuilder oleBuilder = new DocumentBuilder(doc);
        // "Package" progId allows embedding arbitrary files.
        oleBuilder.InsertOleObject("metadata.xml", "Package", false, false, null);

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as PDF/A‑3u.
        doc.Save("output.pdf", saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The PDF/A‑3u file was not created.");

        // Clean up temporary files (optional).
        File.Delete("input.docx");
        File.Delete("metadata.xml");
    }
}
