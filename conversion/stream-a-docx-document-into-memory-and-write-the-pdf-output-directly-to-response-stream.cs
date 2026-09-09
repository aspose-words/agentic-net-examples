using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add some sample text.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOCX content for PDF conversion.");

        // Save the document as a DOCX file locally (bootstrap step).
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Simulate an HTTP response stream using a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();

        // Save the document directly to the simulated response stream in PDF format.
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the stream.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // Optional: display the size of the generated PDF.
        Console.WriteLine($"PDF stream length: {responseStream.Length} bytes.");
    }
}
