using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content for PDF conversion.");

        // Save the sample as a DOC file (required by the workflow).
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Simulate a network response using a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the simulated response.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // Optionally, write the PDF to a file to inspect the result.
        const string outputPath = "output.pdf";
        File.WriteAllBytes(outputPath, responseStream.ToArray());

        // Clean up the temporary DOC file.
        if (File.Exists(inputPath))
            File.Delete(inputPath);
    }
}
