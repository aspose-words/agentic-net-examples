using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple DOCX document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");

        // Save the DOCX to a memory stream.
        using (MemoryStream docxStream = new MemoryStream())
        {
            doc.Save(docxStream, SaveFormat.Docx);
            docxStream.Position = 0; // Reset for reading.

            // Load the document from the DOCX stream.
            Document loadedDoc = new Document(docxStream);

            // Simulate an HTTP response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                // Convert and write the PDF directly to the response stream.
                loadedDoc.Save(responseStream, SaveFormat.Pdf);

                // Validate that the PDF data was written.
                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("The PDF output stream is empty.");
                }

                // (Optional) Reset position if further processing is needed.
                responseStream.Position = 0;
            }
        }
    }
}
