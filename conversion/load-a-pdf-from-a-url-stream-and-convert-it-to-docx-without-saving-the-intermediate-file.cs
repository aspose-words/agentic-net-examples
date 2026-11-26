using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF into a memory stream.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content generated for conversion.");

        using MemoryStream pdfStream = new MemoryStream();
        source.Save(pdfStream, SaveFormat.Pdf);
        pdfStream.Position = 0; // Reset for reading.

        // Load the PDF from the stream (simulating a URL stream).
        Document pdfDocument = new Document(pdfStream);

        // Convert the loaded PDF to DOCX and save to a file.
        const string outputPath = "output.docx";
        pdfDocument.Save(outputPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }
    }
}
