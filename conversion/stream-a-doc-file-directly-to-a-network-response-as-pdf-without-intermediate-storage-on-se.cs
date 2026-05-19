using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank Word document and add sample content.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");

        // Save the document locally as DOC (simulating an existing file).
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Simulate a network response by writing the PDF directly to a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the simulated response stream.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // (Optional) Reset position if further processing is needed.
        responseStream.Position = 0;

        // Example: write the stream to a file to demonstrate the result (not required for the core task).
        const string outputPath = "output.pdf";
        using FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
        responseStream.CopyTo(file);
    }
}
