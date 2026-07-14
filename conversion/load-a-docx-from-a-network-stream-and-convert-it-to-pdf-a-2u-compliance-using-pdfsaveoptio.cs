using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a sample DOCX file.
        const string inputPath = "sample.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document to be converted to PDF/A‑2u.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Simulate receiving the DOCX over a network by loading it into a MemoryStream.
        byte[] fileBytes = File.ReadAllBytes(inputPath);
        using MemoryStream networkStream = new MemoryStream(fileBytes);
        networkStream.Position = 0; // Ensure the stream is at the beginning.

        // Load the document from the simulated network stream.
        Document doc = new Document(networkStream);

        // Configure PDF/A‑2u compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as a PDF/A‑2u file.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑2u file was not created.");

        // Clean up temporary files (optional).
        File.Delete(inputPath);
    }
}
