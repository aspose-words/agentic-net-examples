using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file that will act as the SharePoint document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content generated for conversion.");
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Verify that the input file was created.
        if (!File.Exists(inputPath))
            throw new InvalidOperationException("The input DOCX file was not created.");

        // Step 2: Simulate obtaining the DOCX from a SharePoint stream.
        // Read the file into a byte array and wrap it with a MemoryStream.
        byte[] docxBytes = File.ReadAllBytes(inputPath);
        using (MemoryStream sharePointStream = new MemoryStream(docxBytes))
        {
            // Ensure the stream is positioned at the beginning before loading.
            sharePointStream.Position = 0;

            // Load the document from the simulated SharePoint stream.
            Document docFromStream = new Document(sharePointStream);

            // Step 3: Convert the loaded document to PDF and write it to a response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                docFromStream.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written to the stream.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optional: Save the PDF to a file for verification.
                const string outputPath = "output.pdf";
                responseStream.Position = 0; // Reset before reading.
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(file);
                }

                // Verify that the PDF file exists.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("The output PDF file was not created.");
            }
        }
    }
}
