using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file that would represent a document stored in SharePoint.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document that simulates content retrieved from SharePoint.");
        const string inputPath = "sample.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Verify the input file was created.
        if (!File.Exists(inputPath))
            throw new InvalidOperationException("The sample DOCX file was not created.");

        // Step 2: Simulate obtaining a SharePoint stream containing the DOCX.
        using (MemoryStream sharepointStream = new MemoryStream(File.ReadAllBytes(inputPath)))
        {
            // Ensure the stream is positioned at the beginning before loading.
            sharepointStream.Position = 0;

            // Step 3: Load the document from the simulated SharePoint stream.
            Document docFromSharePoint = new Document(sharepointStream);

            // Step 4: Convert the document to PDF and write it to a simulated response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                docFromSharePoint.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optional: Save the PDF to a file for verification.
                const string outputPath = "output.pdf";
                responseStream.Position = 0;
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(fileStream);
                }

                // Verify that the PDF file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("The output PDF file was not created.");
            }
        }
    }
}
