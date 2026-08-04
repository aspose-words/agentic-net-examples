using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");

        // Save the document to a local DOC file (bootstrap step).
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Simulate a network response by writing the PDF directly to a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Validate that data was written to the simulated response.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // (Optional) Write the PDF to a file for manual inspection.
        // File.WriteAllBytes("output.pdf", responseStream.ToArray());
    }
}
