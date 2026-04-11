using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input DOCX and the resulting PDF.
        const string inputPath = "sample.docx";
        const string outputPath = "sample_PdfA2u.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document locally.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);
        builder.Writeln("Hello Aspose.Words!");
        createDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Simulate loading the DOCX from a network stream.
        // -----------------------------------------------------------------
        byte[] docBytes = File.ReadAllBytes(inputPath);
        using (MemoryStream networkStream = new MemoryStream(docBytes))
        {
            // Ensure the stream is positioned at the beginning before loading.
            networkStream.Position = 0;

            // Load the document from the simulated network stream.
            Document loadedDoc = new Document(networkStream);

            // -----------------------------------------------------------------
            // 3. Convert to PDF/A‑2u using PdfSaveOptions.
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };

            loadedDoc.Save(outputPath, pdfOptions);
        }

        // -----------------------------------------------------------------
        // 4. Validate that the PDF file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF output file was not created.", outputPath);

        FileInfo pdfInfo = new FileInfo(outputPath);
        if (pdfInfo.Length == 0)
            throw new InvalidOperationException("The PDF output file is empty.");

        Console.WriteLine($"Conversion successful. PDF saved to '{outputPath}' ({pdfInfo.Length} bytes).");
    }
}
