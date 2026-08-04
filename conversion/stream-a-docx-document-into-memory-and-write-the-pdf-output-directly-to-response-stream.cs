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
        builder.Writeln("Sample DOCX content for conversion.");

        // Save the sample document to a local DOCX file.
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file that was just created.
        Document loadedDoc = new Document(inputPath);

        // Simulate an HTTP response stream using a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();

        // Save the loaded document as PDF directly into the simulated response stream.
        loadedDoc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the stream.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // Optionally, write the PDF to a file to inspect the result.
        const string outputPath = "output.pdf";
        File.WriteAllBytes(outputPath, responseStream.ToArray());

        // Clean up the temporary input file.
        if (File.Exists(inputPath))
            File.Delete(inputPath);
    }
}
