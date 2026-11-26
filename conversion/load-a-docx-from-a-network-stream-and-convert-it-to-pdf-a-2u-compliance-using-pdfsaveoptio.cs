using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample content for PDF/A-2u conversion.");
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Simulate loading the DOCX from a network stream.
        using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        using (MemoryStream networkStream = new MemoryStream())
        {
            fileStream.CopyTo(networkStream);
            networkStream.Position = 0; // Reset before loading.

            Document docFromStream = new Document(networkStream);

            // Step 3: Convert to PDF/A‑2u using PdfSaveOptions.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };
            const string outputPath = "output.pdf";
            docFromStream.Save(outputPath, pdfOptions);
        }

        // Step 4: Validate that the PDF/A‑2u file was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF/A‑2u file was not created.");
        }
    }
}
