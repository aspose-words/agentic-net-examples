using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputDocPath = "input.doc";
        const string outputPdfPath = "output_pdfa3b.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document and save it as DOC.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for PDF/A‑3b conversion.");
        sourceDoc.Save(inputDocPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Step 2: Load the created DOC file.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocPath);

        // -----------------------------------------------------------------
        // Step 3: Configure PDF save options for PDF/A‑3b compliance and
        //         embed custom XML metadata (archival requirement).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑3b is represented by the PdfA3u compliance level.
            Compliance = PdfCompliance.PdfA3u,

            // Export custom document properties as XMP metadata.
            CustomPropertiesExport = PdfCustomPropertiesExport.Metadata
        };

        // -----------------------------------------------------------------
        // Step 4: Save the document as PDF/A‑3b.
        // -----------------------------------------------------------------
        doc.Save(outputPdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // Step 5: Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"The PDF/A‑3b file '{outputPdfPath}' was not created.");
        }

        // Optional: Clean up intermediate files (not required for the example).
        // File.Delete(inputDocPath);
    }
}
