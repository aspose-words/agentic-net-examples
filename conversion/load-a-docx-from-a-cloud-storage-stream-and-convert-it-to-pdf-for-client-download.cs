using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content generated for conversion.");
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Load the DOCX from a stream (simulating a cloud storage download).
        using (FileStream inputStream = File.OpenRead(inputPath))
        {
            Document loadedDoc = new Document(inputStream);

            // Step 3: Convert the document to PDF and write it to a memory stream
            // (simulating an HTTP response stream for client download).
            using (MemoryStream responseStream = new MemoryStream())
            {
                loadedDoc.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optional: Save the PDF to a file for verification.
                const string outputPath = "output.pdf";
                // Reset the stream position before saving to file.
                responseStream.Position = 0;
                using (FileStream fileStream = File.Create(outputPath))
                {
                    responseStream.CopyTo(fileStream);
                }

                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("Expected output PDF was not created.");
            }
        }
    }
}
