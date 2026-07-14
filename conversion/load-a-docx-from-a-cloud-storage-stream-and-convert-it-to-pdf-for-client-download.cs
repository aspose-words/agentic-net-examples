using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document locally.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content generated for cloud‑storage stream conversion.");

        // Save the sample document as a DOCX file (bootstrap for the input).
        const string inputFileName = "input.docx";
        sourceDoc.Save(inputFileName, SaveFormat.Docx);

        // Step 2: Simulate a cloud storage stream by loading the DOCX file into a MemoryStream.
        using (FileStream fileStream = new FileStream(inputFileName, FileMode.Open, FileAccess.Read))
        using (MemoryStream cloudStream = new MemoryStream())
        {
            fileStream.CopyTo(cloudStream);
            cloudStream.Position = 0; // Reset before reading.

            // Step 3: Load the document from the simulated cloud stream.
            Document docFromCloud = new Document(cloudStream);

            // Step 4: Convert the loaded document to PDF using another MemoryStream.
            using (MemoryStream pdfStream = new MemoryStream())
            {
                docFromCloud.Save(pdfStream, SaveFormat.Pdf);

                // Verify that PDF data was written.
                if (pdfStream.Length == 0)
                    throw new InvalidOperationException("PDF conversion produced an empty stream.");

                // Optional: write the PDF to a file for verification.
                const string outputFileName = "output.pdf";
                pdfStream.Position = 0; // Reset before copying to file.
                using (FileStream outFile = new FileStream(outputFileName, FileMode.Create, FileAccess.Write))
                {
                    pdfStream.CopyTo(outFile);
                }

                // Verify that the output file exists.
                if (!File.Exists(outputFileName))
                    throw new InvalidOperationException("Expected output PDF file was not created.");
            }
        }
    }
}
