using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document and store it in a memory stream
        using (MemoryStream cloudStream = new MemoryStream())
        {
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);
            builder.Writeln("Hello from cloud storage DOCX.");

            // Save the document as DOCX into the simulated cloud stream
            sourceDoc.Save(cloudStream, SaveFormat.Docx);
            cloudStream.Position = 0; // Reset for reading

            // Load the DOCX from the simulated cloud storage stream
            Document loadedDoc = new Document(cloudStream);

            // Convert the loaded document to PDF and write to a simulated client response stream
            using (MemoryStream responseStream = new MemoryStream())
            {
                loadedDoc.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written
                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
                }

                // Optionally, save the PDF to a file for verification
                responseStream.Position = 0;
                using (FileStream file = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(file);
                }

                if (!File.Exists("output.pdf"))
                {
                    throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
