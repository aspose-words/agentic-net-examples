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
        builder.Writeln("Sample content for PDF/A‑3b conversion.");

        // Add a custom document property that contains XML.
        // This property will be exported as XMP metadata.
        doc.CustomDocumentProperties.Add("CustomXml", "<root><info>Sample</info></root>");

        // Configure PDF save options:
        // - PDF/A‑3u compliance (Aspose.Words does not have a direct PdfA3b enum value;
        //   PdfA3u is the closest option that supports PDF/A‑3 features).
        // - Export custom properties as XMP metadata.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            CustomPropertiesExport = PdfCustomPropertiesExport.Metadata
        };

        const string outputPath = "output_pdfa3b.pdf";

        // Save the document as PDF/A‑3b (using PdfA3u compliance) with embedded metadata.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑3b file was not created.");
    }
}
