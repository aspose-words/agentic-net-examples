using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file that would normally reside in SharePoint.
        const string inputPath = "sample.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This document was loaded from a SharePoint stream and converted to PDF.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Simulate obtaining a SharePoint stream containing the DOCX.
        // In a real scenario the stream would come from SharePoint's API.
        using (MemoryStream sharePointStream = new MemoryStream(File.ReadAllBytes(inputPath)))
        {
            // Ensure the stream is positioned at the beginning before loading.
            sharePointStream.Position = 0;

            // Step 3: Load the DOCX from the simulated SharePoint stream.
            Document docFromStream = new Document(sharePointStream);

            // Step 4: Convert the document to PDF and write it to an output stream
            // that represents the HTTP response stream for on‑demand download.
            using (MemoryStream responseStream = new MemoryStream())
            {
                docFromStream.Save(responseStream, SaveFormat.Pdf);

                // Verify that PDF data was written.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optional: write the PDF to a file for local verification.
                const string outputPath = "output.pdf";
                File.WriteAllBytes(outputPath, responseStream.ToArray());

                // Verify that the file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("Expected output PDF file was not created.");
            }
        }
    }
}
