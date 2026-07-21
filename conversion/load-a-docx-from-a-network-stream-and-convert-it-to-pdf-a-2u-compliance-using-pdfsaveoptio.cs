using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple DOCX document in memory.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Sample content for PDF/A‑2u conversion.");

        // Step 2: Save the DOCX to a memory stream to simulate a network stream.
        using (MemoryStream networkStream = new MemoryStream())
        {
            sourceDocument.Save(networkStream, SaveFormat.Docx);
            // Reset the position so the stream can be read from the beginning.
            networkStream.Position = 0;

            // Step 3: Load the document from the simulated network stream.
            Document loadedDocument = new Document(networkStream);

            // Step 4: Configure PDF save options for PDF/A‑2u compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };

            // Step 5: Save the loaded document as a PDF/A‑2u file.
            const string outputPdfPath = "output.pdf";
            loadedDocument.Save(outputPdfPath, pdfOptions);

            // Step 6: Validate that the PDF file was created.
            if (!File.Exists(outputPdfPath))
                throw new InvalidOperationException("The PDF/A‑2u output file was not created.");

            // Optional: Verify that the file has content.
            FileInfo fileInfo = new FileInfo(outputPdfPath);
            if (fileInfo.Length == 0)
                throw new InvalidOperationException("The PDF/A‑2u output file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document successfully converted to PDF/A‑2u.");
    }
}
