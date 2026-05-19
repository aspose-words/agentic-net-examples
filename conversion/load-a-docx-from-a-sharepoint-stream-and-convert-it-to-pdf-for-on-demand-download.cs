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
        builder.Writeln("Sample content loaded from a SharePoint stream.");
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Simulate a SharePoint stream by reading the DOCX into a MemoryStream.
        using (FileStream fileStream = File.OpenRead(inputPath))
        using (MemoryStream sharepointStream = new MemoryStream())
        {
            fileStream.CopyTo(sharepointStream);
            sharepointStream.Position = 0; // Reset for reading.

            // Step 3: Load the document from the simulated SharePoint stream.
            Document doc = new Document(sharepointStream);

            // Step 4: Convert the document to PDF and write to a simulated response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optional: Save the PDF to a file for verification.
                const string outputPath = "output.pdf";
                responseStream.Position = 0;
                using (FileStream outFile = File.Create(outputPath))
                {
                    responseStream.CopyTo(outFile);
                }

                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("Expected output PDF was not created.");
            }
        }
    }
}
