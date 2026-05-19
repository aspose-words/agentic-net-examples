using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample content for cloud storage conversion.");

        // Save the sample DOCX to a local file (simulating a file stored in the cloud).
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX into a memory stream to simulate downloading it from cloud storage.
        byte[] docBytes = File.ReadAllBytes(inputPath);
        using (MemoryStream cloudStream = new MemoryStream(docBytes))
        {
            // Ensure the stream position is at the beginning before loading.
            cloudStream.Position = 0;

            // Load the document from the simulated cloud stream.
            Document doc = new Document(cloudStream);

            // Convert the document to PDF and write it to a response stream (simulating client download).
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);

                // Verify that PDF data was written to the response stream.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optionally, write the PDF to a local file for verification.
                const string outputPath = "output.pdf";
                File.WriteAllBytes(outputPath, responseStream.ToArray());

                // Verify that the output file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("Expected output PDF was not created.");
            }
        }
    }
}
