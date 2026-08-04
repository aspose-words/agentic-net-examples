using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file that will act as the source document.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Sample content for PDF/A‑2u conversion.");
        const string inputPath = "input.docx";
        sourceDocument.Save(inputPath, SaveFormat.Docx);

        // Simulate receiving the DOCX over a network by loading it from a memory stream.
        byte[] fileBytes = File.ReadAllBytes(inputPath);
        using (MemoryStream networkStream = new MemoryStream(fileBytes))
        {
            // Reset the stream position before loading.
            networkStream.Position = 0;

            // Load the document from the simulated network stream.
            Document loadedDocument = new Document(networkStream);

            // Configure PDF save options for PDF/A‑2u compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.Compliance = PdfCompliance.PdfA2u;

            // Save the document as a PDF/A‑2u file.
            const string outputPath = "output.pdf";
            loadedDocument.Save(outputPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The PDF/A‑2u output file was not created.");
        }

        // Clean up temporary files (optional).
        File.Delete(inputPath);
    }
}
