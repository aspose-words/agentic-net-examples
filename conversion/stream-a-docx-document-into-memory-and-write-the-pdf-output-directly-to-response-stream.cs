using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample DOCX document created in memory.");

        // Save the document to a memory stream in DOCX format.
        using (MemoryStream docxStream = new MemoryStream())
        {
            source.Save(docxStream, SaveFormat.Docx);
            docxStream.Position = 0; // Reset for reading.

            // Load the document from the DOCX stream.
            Document doc = new Document(docxStream);

            // Simulate an HTTP response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                // Convert the document to PDF and write directly to the response stream.
                doc.Save(responseStream, SaveFormat.Pdf);

                // Verify that PDF data was written.
                if (responseStream.Length == 0)
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

                // Optionally write the PDF to a file for inspection.
                responseStream.Position = 0;
                using (FileStream file = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(file);
                }

                if (!File.Exists("output.pdf"))
                    throw new InvalidOperationException("Expected output PDF file was not created.");
            }
        }
    }
}
