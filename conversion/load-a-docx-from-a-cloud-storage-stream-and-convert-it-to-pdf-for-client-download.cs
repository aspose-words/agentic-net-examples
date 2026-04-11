using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define temporary file names.
        const string docxFileName = "Sample.docx";
        const string pdfFileName = "Result.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello from Aspose.Words!");
        sampleDoc.Save(docxFileName, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Simulate loading the DOCX from a cloud storage stream.
        // -----------------------------------------------------------------
        using (FileStream fileStream = File.OpenRead(docxFileName))
        using (MemoryStream cloudStream = new MemoryStream())
        {
            fileStream.CopyTo(cloudStream);
            cloudStream.Position = 0; // Reset before reading.

            // Load the document from the stream.
            Document loadedDoc = new Document(cloudStream);

            // -----------------------------------------------------------------
            // 3. Convert the loaded document to PDF and write to an output stream.
            // -----------------------------------------------------------------
            using (MemoryStream pdfStream = new MemoryStream())
            {
                // Simple conversion using the native Save API.
                loadedDoc.Save(pdfStream, SaveFormat.Pdf);

                // Verify that the conversion produced data.
                if (pdfStream.Length == 0)
                    throw new InvalidOperationException("PDF conversion failed: output stream is empty.");

                // Optionally, write the PDF to a file for verification.
                File.WriteAllBytes(pdfFileName, pdfStream.ToArray());
            }
        }

        // Clean up the temporary DOCX file.
        if (File.Exists(docxFileName))
            File.Delete(docxFileName);
    }
}
