using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document locally.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content generated for conversion.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Simulate loading the DOCX from a cloud storage stream.
        using (MemoryStream cloudStream = new MemoryStream())
        {
            using (FileStream fileStream = new FileStream("input.docx", FileMode.Open, FileAccess.Read))
            {
                fileStream.CopyTo(cloudStream);
            }
            cloudStream.Position = 0; // Reset before reading.

            // Load the document from the simulated cloud stream.
            Document doc = new Document(cloudStream);

            // Convert the document to PDF and write to a simulated response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                doc.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written.
                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
                }

                // Optional: write the PDF to a file for verification.
                responseStream.Position = 0;
                using (FileStream outputFile = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(outputFile);
                }

                if (!File.Exists("output.pdf"))
                {
                    throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
