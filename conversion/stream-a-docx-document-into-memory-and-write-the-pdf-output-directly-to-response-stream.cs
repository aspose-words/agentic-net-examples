using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample DOCX document created for conversion.");

        // Save the DOCX to a memory stream.
        using (MemoryStream docxStream = new MemoryStream())
        {
            sourceDoc.Save(docxStream, SaveFormat.Docx);
            // Reset the stream position before loading.
            docxStream.Position = 0;

            // Load the document from the DOCX memory stream.
            Document loadedDoc = new Document(docxStream);

            // Simulate a response stream and save the PDF directly into it.
            using (MemoryStream responseStream = new MemoryStream())
            {
                loadedDoc.Save(responseStream, SaveFormat.Pdf);

                // Verify that PDF data was written.
                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
                }

                // Optionally, you could write the PDF to a file for verification.
                // File.WriteAllBytes("output.pdf", responseStream.ToArray());
            }
        }
    }
}
