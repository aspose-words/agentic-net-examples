using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document from the file.
        Document doc = new Document("input.docx");

        // Simulate an HTTP response stream using MemoryStream.
        using MemoryStream responseStream = new MemoryStream();

        // Save the document as PDF directly into the response stream.
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the stream.
        if (responseStream.Length == 0)
        {
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
        }

        // (Optional) Reset position if further processing is needed.
        responseStream.Position = 0;

        // Example: write the PDF to a file to confirm the conversion succeeded.
        // This step is not required for the streaming scenario but demonstrates the output.
        using FileStream file = new FileStream("output.pdf", FileMode.Create, FileAccess.Write);
        responseStream.CopyTo(file);
    }
}
