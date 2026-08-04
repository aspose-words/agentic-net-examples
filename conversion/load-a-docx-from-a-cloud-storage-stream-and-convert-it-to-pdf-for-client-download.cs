using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document locally.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content.");
        const string inputFileName = "input.docx";
        sourceDoc.Save(inputFileName, SaveFormat.Docx);

        // Step 2: Simulate loading the DOCX from a cloud storage stream.
        using (FileStream fileStream = File.OpenRead(inputFileName))
        using (MemoryStream cloudStream = new MemoryStream())
        {
            fileStream.CopyTo(cloudStream);
            cloudStream.Position = 0; // Reset for reading.

            // Step 3: Load the document from the simulated cloud stream.
            Document doc = new Document(cloudStream);

            // Step 4: Convert the document to PDF and write to a simulated response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optional: Save the PDF to a file for verification.
                const string outputFileName = "output.pdf";
                responseStream.Position = 0;
                using (FileStream outFile = File.Create(outputFileName))
                {
                    responseStream.CopyTo(outFile);
                }

                // Verify that the output file exists.
                if (!File.Exists(outputFileName))
                    throw new InvalidOperationException("Expected output PDF was not created.");
            }
        }
    }
}
