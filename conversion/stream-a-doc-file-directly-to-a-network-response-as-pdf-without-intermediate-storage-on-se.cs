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

        // Simulate a network response stream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Validate that PDF data was written.
        if (responseStream.Length == 0)
        {
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
        }

        // (Optional) Reset position if further processing is needed.
        responseStream.Position = 0;

        // Example: write the PDF to a file to verify the conversion succeeded.
        const string outputPath = "output.pdf";
        using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            responseStream.CopyTo(file);
        }

        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
