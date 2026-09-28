using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for PDF/A‑2u conversion.");
        const string inputPath = "sample.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Simulate loading the DOCX from a network stream.
        using (MemoryStream networkStream = new MemoryStream())
        {
            // Write the DOCX file content to the simulated network stream.
            using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                fileStream.CopyTo(networkStream);
            }

            // Reset the stream position before loading.
            networkStream.Position = 0;

            // Load the document from the stream.
            Document loadedDoc = new Document(networkStream);

            // Configure PDF/A‑2u compliance options.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };

            // Save the converted PDF/A‑2u document.
            const string outputPath = "output.pdf";
            loadedDoc.Save(outputPath, saveOptions);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Expected output PDF/A‑2u file was not created.");
        }

        // Clean up temporary input file.
        if (File.Exists(inputPath))
            File.Delete(inputPath);
    }
}
