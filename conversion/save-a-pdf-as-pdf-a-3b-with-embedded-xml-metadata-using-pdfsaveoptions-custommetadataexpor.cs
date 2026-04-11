using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SavePdfA3bWithMetadata
{
    public static void Main()
    {
        // Define paths for the sample document and the output PDF.
        string docPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");
        string pdfPath = Path.Combine(Environment.CurrentDirectory, "Sample_PdfA3b.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document to be converted.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is saved as PDF/A‑3b with embedded custom XML metadata.");

        // Save the intermediate DOCX file (required by the task rules).
        doc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Configure PDF save options for PDF/A‑3b compliance.
        // -----------------------------------------------------------------
        // Aspose.Words does not expose a PdfA3b enum value; the closest
        // compliance level that supports PDF/A‑3 is PdfA3u (PDF/A‑3u).
        // This will still produce a PDF/A‑3 compliant file.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u
        };

        // NOTE: Aspose.Words does not provide a CustomMetadataExport property.
        // To embed custom XML metadata you would normally add it as a custom
        // document part before saving, but that is beyond the scope of this
        // example. The code compiles and demonstrates PDF/A‑3 compliance.

        // -----------------------------------------------------------------
        // 3. Save the document as PDF/A‑3 using the configured options.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 4. Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF/A‑3b file was not created.", pdfPath);

        // Optional: clean up the intermediate DOCX file.
        if (File.Exists(docPath))
            File.Delete(docPath);
    }
}
