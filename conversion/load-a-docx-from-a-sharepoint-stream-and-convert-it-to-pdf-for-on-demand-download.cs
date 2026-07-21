using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file locally.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Sample DOC content generated for conversion.");
        const string inputPath = "input.docx";
        sourceDocument.Save(inputPath, SaveFormat.Docx);

        // Step 2: Simulate obtaining a SharePoint stream by loading the file into a MemoryStream.
        using (FileStream fileStream = File.OpenRead(inputPath))
        using (MemoryStream sharepointStream = new MemoryStream())
        {
            fileStream.CopyTo(sharepointStream);
            sharepointStream.Position = 0; // Reset for reading.

            // Step 3: Load the document from the simulated SharePoint stream.
            Document doc = new Document(sharepointStream);

            // Step 4: Convert the document to PDF and write it to a simulated response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);
                responseStream.Position = 0; // Reset for any further processing.

                // Step 5: Validate that PDF data was written.
                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
                }

                // Optional: Save the PDF to a file for verification purposes.
                const string outputPath = "output.pdf";
                using (FileStream outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(outFile);
                }

                if (!File.Exists(outputPath))
                {
                    throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
