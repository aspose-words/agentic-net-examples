using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOCX content.");

        // Save the document as a local DOCX file to simulate an existing input file.
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from the file.
        Document doc = new Document(inputPath);

        // Simulate an HTTP response stream using a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();

        // Save the document as PDF directly into the simulated response stream.
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the stream.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // (Optional) Write the PDF to a file for manual inspection.
        // File.WriteAllBytes("output.pdf", responseStream.ToArray());
    }
}
