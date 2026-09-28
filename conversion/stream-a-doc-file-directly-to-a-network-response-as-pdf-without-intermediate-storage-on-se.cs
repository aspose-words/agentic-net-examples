using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content for PDF conversion.");
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Simulate a network response using a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Verify that PDF data was written to the simulated response stream.
        if (responseStream.Length == 0)
        {
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
        }

        // Reset the stream position before further processing.
        responseStream.Position = 0;

        // Optional: write the PDF to a file to demonstrate the result.
        const string outputPath = "output.pdf";
        using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            responseStream.CopyTo(file);
        }

        // Validate that the output PDF file was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
