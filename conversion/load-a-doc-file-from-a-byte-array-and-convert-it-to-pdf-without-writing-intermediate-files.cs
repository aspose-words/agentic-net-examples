using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOC document in memory.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This document is created in memory and will be loaded from a byte array.");

        // Step 2: Save the sample document to a MemoryStream in DOC format.
        byte[] docBytes;
        using (MemoryStream docStream = new MemoryStream())
        {
            sampleDoc.Save(docStream, SaveFormat.Doc);
            docBytes = docStream.ToArray(); // Capture the byte array.
        }

        // Step 3: Load a Document from the byte array.
        Document loadedDoc;
        using (MemoryStream loadStream = new MemoryStream(docBytes))
        {
            loadedDoc = new Document(loadStream); // Format is detected automatically.
        }

        // Step 4: Convert the loaded document to PDF using another MemoryStream.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            loadedDoc.Save(pdfStream, SaveFormat.Pdf);

            // Validation: ensure the PDF stream contains data.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("PDF conversion failed: output stream is empty.");

            // Optional: write the PDF to a temporary file for visual verification.
            string outputPath = Path.Combine(Path.GetTempPath(), "ConvertedFromByteArray.pdf");
            File.WriteAllBytes(outputPath, pdfStream.ToArray());

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("PDF file was not created.", outputPath);
        }
    }
}
