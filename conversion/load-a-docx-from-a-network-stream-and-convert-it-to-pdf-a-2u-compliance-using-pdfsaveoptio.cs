using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary DOCX and the final PDF/A‑2u file.
        const string inputDocxPath = "sample.docx";
        const string outputPdfPath = "sample_pdfa2u.pdf";

        // 1. Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for PDF/A‑2u conversion.");
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // 2. Simulate a network download by reading the DOCX into a MemoryStream.
        byte[] docBytes = File.ReadAllBytes(inputDocxPath);
        using (MemoryStream networkStream = new MemoryStream(docBytes))
        {
            // Ensure the stream is positioned at the start before loading.
            networkStream.Position = 0;

            // 3. Load the document from the simulated network stream.
            Document loadedDoc = new Document(networkStream);

            // 4. Set up PDF/A‑2u compliance options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };

            // 5. Save the document as a PDF/A‑2u file.
            loadedDoc.Save(outputPdfPath, pdfOptions);
        }

        // Verify that the PDF/A‑2u file was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("PDF/A‑2u output file was not created.");
        }
    }
}
