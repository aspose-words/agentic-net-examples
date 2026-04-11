using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the generated files.
        const string xmlFilePath = "metadata.xml";
        const string pdfOutputPath = "DocumentPdfA3u.pdf";

        // Create a simple XML file that will be embedded.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<metadata>
    <author>John Doe</author>
    <description>Sample XML metadata embedded in PDF/A‑3u.</description>
</metadata>";
        File.WriteAllText(xmlFilePath, xmlContent);

        // Create a blank Word document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an embedded XML file for enhanced metadata.");

        // Embed the XML file as an OLE object. The progId "Package" works for generic files.
        builder.InsertOleObject(xmlFilePath, "Package", false, true, null);

        // Configure PDF save options for PDF/A‑3u compliance and enable attachment embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as PDF/A‑3u.
        doc.Save(pdfOutputPath, saveOptions);

        // Validate that the PDF file was created and is not empty.
        if (!File.Exists(pdfOutputPath) || new FileInfo(pdfOutputPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the PDF/A‑3u document.");
        }

        // Clean up the temporary XML file (optional).
        if (File.Exists(xmlFilePath))
        {
            File.Delete(xmlFilePath);
        }
    }
}
