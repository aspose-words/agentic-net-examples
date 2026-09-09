using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");

        // Save the document locally as DOC (simulating an existing file).
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Simulate a network response using a memory stream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // Reset the stream position and optionally write the PDF to a file for inspection.
        responseStream.Position = 0;
        using FileStream file = new FileStream("output.pdf", FileMode.Create, FileAccess.Write);
        responseStream.CopyTo(file);
    }
}
