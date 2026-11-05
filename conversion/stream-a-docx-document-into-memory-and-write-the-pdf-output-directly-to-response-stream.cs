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
        builder.Writeln("This is a sample DOCX document created for conversion.");

        // Save the document as DOCX to a local file.
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from the file.
        Document doc = new Document(inputPath);

        // Simulate an HTTP response stream using MemoryStream.
        using MemoryStream responseStream = new MemoryStream();

        // Save the document as PDF directly into the response stream.
        doc.Save(responseStream, SaveFormat.Pdf);

        // Validate that PDF data was written to the stream.
        if (responseStream.Length == 0)
        {
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
        }

        // (Optional) Write the PDF to a file for verification.
        const string outputPath = "output.pdf";
        File.WriteAllBytes(outputPath, responseStream.ToArray());

        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF file was not created.");
        }

        // Clean up the temporary input file.
        if (File.Exists(inputPath))
        {
            File.Delete(inputPath);
        }
    }
}
