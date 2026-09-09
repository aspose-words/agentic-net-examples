using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file that represents a document stored in SharePoint.
        const string inputFileName = "input.docx";
        if (!File.Exists(inputFileName))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document that would be stored in SharePoint.");
            sampleDoc.Save(inputFileName, SaveFormat.Docx);
        }

        // Step 2: Simulate obtaining a stream from SharePoint.
        using (FileStream sharePointStream = new FileStream(inputFileName, FileMode.Open, FileAccess.Read))
        {
            // Step 3: Load the document from the simulated SharePoint stream.
            Document doc = new Document(sharePointStream);

            // Step 4: Convert the document to PDF and write it to a response‑like stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);

                // Ensure the stream contains data.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Reset position before any further read/copy operations.
                responseStream.Position = 0;

                // Optional: Save the PDF to a file for verification.
                const string outputFileName = "output.pdf";
                using (FileStream fileOut = new FileStream(outputFileName, FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(fileOut);
                }

                if (!File.Exists(outputFileName))
                    throw new InvalidOperationException("Expected output PDF was not created.");
            }
        }
    }
}
