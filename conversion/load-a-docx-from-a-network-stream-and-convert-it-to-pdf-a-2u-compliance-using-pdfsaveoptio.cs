using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file locally.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document created for PDF/A‑2u conversion.");
        const string inputPath = "sample_input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Simulate loading the DOCX from a network stream.
        byte[] docBytes = File.ReadAllBytes(inputPath);
        using (MemoryStream networkStream = new MemoryStream(docBytes))
        {
            // Ensure the stream is positioned at the beginning before loading.
            networkStream.Position = 0;
            Document loadedDoc = new Document(networkStream);

            // Step 3: Convert the document to PDF/A‑2u compliant PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };
            const string outputPath = "converted_output.pdf";
            loadedDoc.Save(outputPath, pdfOptions);

            // Step 4: Validate that the PDF file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The PDF/A‑2u output file was not created.");
        }

        // Clean up temporary files (optional).
        // File.Delete(inputPath);
        // File.Delete("converted_output.pdf");
    }
}
